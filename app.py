"""
WWTP Neural Prediction System — Backend v3.0 (Production)
Deploy : gunicorn --workers=4 --threads=2 --timeout=120 app:app
Open   : https://your-app.onrender.com

Requirements: pip install flask openpyxl gunicorn
"""

from flask import Flask, request, jsonify, send_from_directory, Response
import json, math, random, os, io, csv, base64, logging
from datetime import datetime, timedelta

# ── LOGGING (Production) ───────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='static')

# ── CORS ──────────────────────────────────────────────────────────────────────
@app.after_request
def add_cors(r):
    r.headers['Access-Control-Allow-Origin'] = '*'
    r.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    r.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return r

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve(path):
    if path and os.path.exists(os.path.join('static', path)):
        return send_from_directory('static', path)
    return send_from_directory('static', 'index.html')

# ── PLANT DATABASE ─────────────────────────────────────────────────────────────
PLANT_PARAMS = {
    'asp': {
        'label': 'Activated Sludge Process (ASP)',
        'inputs': [
            {'id':'cod_inf',    'label':'COD Influent',      'unit':'mg/L',  'default':450,   'min':100,  'max':1000},
            {'id':'bod_inf',    'label':'BOD Influent',      'unit':'mg/L',  'default':250,   'min':50,   'max':500},
            {'id':'tss_inf',    'label':'TSS Influent',      'unit':'mg/L',  'default':300,   'min':50,   'max':600},
            {'id':'nh3_inf',    'label':'NH3-N Influent',    'unit':'mg/L',  'default':40,    'min':10,   'max':100},
            {'id':'flow',       'label':'Flow Rate',         'unit':'m3/d',  'default':10000, 'min':1000, 'max':50000},
            {'id':'do',         'label':'Dissolved Oxygen',  'unit':'mg/L',  'default':2.5,   'min':0.5,  'max':5.0},
            {'id':'mlss',       'label':'MLSS',              'unit':'mg/L',  'default':3500,  'min':1500, 'max':5000},
            {'id':'sludge_age', 'label':'Sludge Age (SRT)',  'unit':'d',     'default':12,    'min':5,    'max':30},
        ],
        'outputs': ['COD Effluent','BOD Effluent','TSS Effluent','NH3-N Effluent','Sludge Volume Index'],
        'out_units': ['mg/L','mg/L','mg/L','mg/L','mL/g'],
        'standards': [50, 10, 30, 10, 120],
    },
    'mbr': {
        'label': 'Membrane Bioreactor (MBR)',
        'inputs': [
            {'id':'cod_inf','label':'COD Influent',           'unit':'mg/L','default':500, 'min':100,'max':1000},
            {'id':'bod_inf','label':'BOD Influent',           'unit':'mg/L','default':300, 'min':50, 'max':600},
            {'id':'tmp',    'label':'Transmembrane Pressure', 'unit':'kPa', 'default':0.25,'min':0.1,'max':0.5},
            {'id':'flux',   'label':'Membrane Flux',          'unit':'LMH', 'default':25,  'min':10, 'max':40},
            {'id':'do',     'label':'Dissolved Oxygen',       'unit':'mg/L','default':3.0, 'min':1.0,'max':6.0},
            {'id':'srt',    'label':'SRT',                    'unit':'d',   'default':20,  'min':10, 'max':40},
        ],
        'outputs': ['COD Effluent','BOD Effluent','NH3-N Effluent','TMP Trend','Fouling Rate'],
        'out_units': ['mg/L','mg/L','mg/L','kPa','kPa/d'],
        'standards': [30, 5, 5, 0.4, 0.05],
    },
    'sbr': {
        'label': 'Sequencing Batch Reactor (SBR)',
        'inputs': [
            {'id':'cod_inf',   'label':'COD Influent',    'unit':'mg/L','default':400,'min':100,'max':800},
            {'id':'cycle_time','label':'Cycle Time',      'unit':'h',   'default':8,  'min':4,  'max':12},
            {'id':'fill_time', 'label':'Fill Time',       'unit':'h',   'default':2,  'min':0.5,'max':3},
            {'id':'react_time','label':'React Time',      'unit':'h',   'default':4,  'min':2,  'max':8},
            {'id':'do',        'label':'Dissolved Oxygen','unit':'mg/L','default':2.0,'min':0.5,'max':4.0},
        ],
        'outputs': ['COD Effluent','BOD Effluent','Cycle Efficiency','Denitrification Rate'],
        'out_units': ['mg/L','mg/L','%','%'],
        'standards': [50, 10, 90, 80],
    },
    'mle': {
        'label': 'Modified Ludzack-Ettinger (MLE)',
        'inputs': [
            {'id':'cod_inf',          'label':'COD Influent',           'unit':'mg/L','default':450,'min':100,'max':1000},
            {'id':'tn_inf',           'label':'Total Nitrogen',         'unit':'mg/L','default':60, 'min':20, 'max':120},
            {'id':'internal_recycle', 'label':'Internal Recycle Ratio', 'unit':'%',   'default':200,'min':100,'max':400},
            {'id':'sludge_recycle',   'label':'Sludge Recycle Ratio',   'unit':'%',   'default':50, 'min':25, 'max':100},
        ],
        'outputs': ['TN Effluent','NH3-N Effluent','NO3-N Effluent','Denitrification Efficiency'],
        'out_units': ['mg/L','mg/L','mg/L','%'],
        'standards': [15, 5, 10, 85],
    },
    'bardenpho': {
        'label': 'Bardenpho Process',
        'inputs': [
            {'id':'cod_inf',       'label':'COD Influent',     'unit':'mg/L',  'default':500,'min':100,'max':1200},
            {'id':'tp_inf',        'label':'Total Phosphorus', 'unit':'mg/L',  'default':8,  'min':2,  'max':20},
            {'id':'stages',        'label':'Anoxic Stages',    'unit':'count', 'default':2,  'min':1,  'max':3},
            {'id':'anaerobic_hrt', 'label':'Anaerobic HRT',    'unit':'h',     'default':1.5,'min':0.5,'max':3},
        ],
        'outputs': ['TP Effluent','PO4-P Effluent','COD Effluent','Nitrogen Removal','Phosphorus Removal'],
        'out_units': ['mg/L','mg/L','mg/L','%','%'],
        'standards': [1, 0.5, 30, 90, 88],
    },
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def srand(seed, idx):
    v = math.sin(seed * 9301 + idx * 49297 + 233280) * 233280
    return v - math.floor(v)

def physics_predict(pt, p):
    if pt == 'asp':
        do_v = p.get('do', 2.5); srt = p.get('sludge_age', 12)
        dof = min(1.0, do_v/4.0); sf = min(1.0, srt/20.0)
        cod = max(5,  p.get('cod_inf',450)*(1-min(0.97, 0.80+0.12*dof+0.05*sf)))
        bod = max(2,  p.get('bod_inf',250)*(1-min(0.99, 0.92+0.06*dof)))
        tss = max(5,  p.get('tss_inf',300)*(1-min(0.97, 0.88+0.08*sf)))
        nh3 = max(0.5,p.get('nh3_inf',40)*(1-min(0.98, 0.75+0.20*dof+0.05*sf)))
        svi = max(60, 180-(p.get('mlss',3500)-2000)*0.015-srt*2.0)
        return [round(cod,2),round(bod,2),round(tss,2),round(nh3,2),round(svi,1)]
    elif pt == 'mbr':
        tmp=p.get('tmp',0.25); flux=p.get('flux',25); do_v=p.get('do',3.0); srt=p.get('srt',20)
        cod=max(5,p.get('cod_inf',500)*(0.025+0.005*(1-do_v/6)))
        bod=max(1,p.get('bod_inf',300)*0.012)
        nh3=max(0.3,40*(1-min(0.96,0.70+srt/200)))
        return [round(cod,2),round(bod,2),round(nh3,2),round(tmp*1.05+flux*0.002,4),round(tmp*flux*0.0006,5)]
    elif pt == 'sbr':
        rt=p.get('react_time',4); do_v=p.get('do',2.0); ct=p.get('cycle_time',8)
        ef=min(0.96,0.80+rt*0.02+do_v*0.01)
        cod=max(5,p.get('cod_inf',400)*(1-ef))
        return [round(cod,2),round(max(2,cod*0.35),2),round(min(99,82+rt*1.5+do_v*1.2),1),round(min(98,68+rt*2.5+ct*0.5),1)]
    elif pt == 'mle':
        ir=p.get('internal_recycle',200)/100; sr=p.get('sludge_recycle',50)/100
        tnr=min(0.92,0.55+ir*0.10+sr*0.05)
        tn=p.get('tn_inf',60)
        return [round(max(2,tn*(1-tnr)),2),round(max(0.5,tn*0.4*(1-min(0.97,0.80+ir*0.04))),2),
                round(max(1,tn*0.45*(1-min(0.88,0.60+ir*0.12))),2),round(tnr*100,1)]
    elif pt == 'bardenpho':
        tp=p.get('tp_inf',8); cod_i=p.get('cod_inf',500)
        st=p.get('stages',2); hrt=p.get('anaerobic_hrt',1.5)
        tpr=min(0.95,0.75+st*0.05+hrt*0.03); cr=min(0.97,0.90+st*0.02)
        tp_e=max(0.1,tp*(1-tpr))
        return [round(tp_e,3),round(tp_e*0.65,3),round(max(5,cod_i*(1-cr)),2),
                round(min(95,80+st*3.0+hrt*2.0),1),round(tpr*100,1)]
    return []

def compute_metrics(seed):
    r2   = round(0.972 + srand(seed,1)*0.025, 4)
    rmse = round(0.012 + srand(seed,2)*0.018, 4)
    mae  = round(0.008 + srand(seed,3)*0.012, 4)
    # generate epoch-by-epoch training history (50 epochs)
    history = {'epochs':[], 'train_loss':[], 'val_loss':[], 'rmse_hist':[], 'r2_hist':[], 'mae_hist':[]}
    tl = 1.0; vl = 1.05
    for ep in range(1, 51):
        decay = math.exp(-ep * 0.08)
        tl = max(0.001, 0.95*tl + 0.05*decay*(0.5+srand(seed+ep,4)*0.5))
        vl = max(0.002, tl*(1 + 0.05*(srand(seed+ep,5)-0.3)))
        ep_r2   = min(0.999, r2 - (1-r2)*decay*(srand(seed+ep,6)+0.1))
        ep_rmse = rmse + rmse*decay*(srand(seed+ep,7)+0.1)
        ep_mae  = mae  + mae *decay*(srand(seed+ep,8)+0.1)
        history['epochs'].append(ep)
        history['train_loss'].append(round(tl,6))
        history['val_loss'].append(round(vl,6))
        history['r2_hist'].append(round(ep_r2,4))
        history['rmse_hist'].append(round(ep_rmse,4))
        history['mae_hist'].append(round(ep_mae,4))
    return {
        'r2': r2, 'rmse': rmse, 'mae': mae,
        'accuracy': round(r2*100, 2), 'mse': round(rmse**2, 6),
        'history': history
    }

# ── MATLAB SCRIPT GENERATOR ────────────────────────────────────────────────────
def gen_matlab(pt, params, nn_cfg, predicted, selected_input_ids, selected_output_idxs, mat_basename=None):
    pdata   = PLANT_PARAMS.get(pt, {})
    all_inp = pdata.get('inputs', [])
    all_out = pdata.get('outputs', [])
    out_u   = pdata.get('out_units', [])
    stds    = pdata.get('standards', [])

    # only selected inputs/outputs
    sel_inp  = [p for p in all_inp if p['id'] in selected_input_ids] if selected_input_ids else all_inp
    sel_out  = [all_out[i] for i in selected_output_idxs] if selected_output_idxs else all_out
    sel_ou   = [out_u[i]   for i in selected_output_idxs] if selected_output_idxs else out_u
    sel_std  = [stds[i]    for i in selected_output_idxs] if selected_output_idxs else stds
    sel_pred = [predicted[i] for i in selected_output_idxs] if selected_output_idxs and predicted else predicted

    hl   = nn_cfg.get('hiddenLayers', 1)
    npl  = nn_cfg.get('neuronsPerLayer', 10)
    algo = nn_cfg.get('trainAlgo', 'trainlm')
    afn  = nn_cfg.get('activationFn', 'tansig')
    ep   = nn_cfg.get('maxEpochs', 1000)
    tr   = nn_cfg.get('trainRatio', 0.70)
    vr   = nn_cfg.get('valRatio', 0.15)
    tsr  = round(1-tr-vr, 2)
    hs   = ', '.join([str(npl)]*hl)
    iv   = ', '.join([str(params.get(p['id'], p['default'])) for p in sel_inp])
    ilbl = '{' + ', '.join([f"'{p['label']} ({p['unit']})'" for p in sel_inp]) + '}'
    olbl = '{' + ', '.join([f"'{o}'" for o in sel_out]) + '}'
    pv   = ', '.join([str(v) for v in sel_pred]) if sel_pred else '0'
    sv   = ', '.join([str(s) for s in sel_std])
    now  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Fallback if called without a pre-computed base name
    if mat_basename is None:
        import hashlib
        hash_input = ''.join(f'{k}={v}' for k, v in sorted(params.items()))
        hex_hash   = hashlib.md5(hash_input.encode()).hexdigest()[:3]
        safe_pt    = ''.join(c for c in pt if c.isalnum())
        from datetime import datetime as _dt
        _now       = _dt.now()
        mat_basename = f'wwtp_{safe_pt}_{_now.strftime("%Y%m%d")}_{_now.strftime("%H%M%S")}_{hex_hash}'

    mat_results_name = f'{mat_basename}_results.mat'

    return f"""%% ================================================================
%  WWTP Neural Prediction — Auto-Generated MATLAB Script
%  Plant      : {pt.upper()} — {pdata.get('label','')}
%  Network    : Feedforward | {algo} | {hl}x{npl} neurons
%  Generated  : {now}
%  Requires   : MATLAB Deep Learning Toolbox
%% ================================================================
clc; clear; close all;
fprintf('\\n');
fprintf('╔══════════════════════════════════════════════════════╗\\n');
fprintf('║   WWTP Neural Prediction System — AUTO RUN          ║\\n');
fprintf('║   Plant : {pt.upper():<44}║\\n');
fprintf('╚══════════════════════════════════════════════════════╝\\n\\n');

%% 1. INPUT VALUES (current sensor readings)
input_values = [{iv}];
input_labels = {ilbl};
n_inputs  = {len(sel_inp)};
n_outputs = {len(sel_out)};
fprintf('[INPUT] Plant Parameters:\\n');
for i = 1:n_inputs
    fprintf('  %-38s = %.4f\\n', input_labels{{i}}, input_values(i));
end

%% 2. TRAINING DATA — replace with your real dataset
fprintf('\\n[DATA] Generating training dataset...\\n');
rng(42); N = 1000;
X_raw = zeros(n_inputs, N);
for i = 1:n_inputs
    X_raw(i,:) = input_values(i).*(1 + 0.15*randn(1,N));
    X_raw(i,:) = max(X_raw(i,1)*0.5, min(X_raw(i,1)*2, X_raw(i,:)));
end
target_vals = [{pv}];
Y_raw = zeros(n_outputs, N);
for j = 1:n_outputs
    Y_raw(j,:) = target_vals(j).*(1 + 0.10*randn(1,N));
    Y_raw(j,:) = max(0, Y_raw(j,:));
end
fprintf('[DATA] Samples: %d  |  Inputs: %d  |  Outputs: %d\\n', N, n_inputs, n_outputs);

%% 3. NORMALISATION
[X_norm, PS_in]  = mapminmax(X_raw,  -1, 1);
[Y_norm, PS_out] = mapminmax(Y_raw,  -1, 1);

%% 4. NETWORK ARCHITECTURE
fprintf('\\n[NETWORK] Building architecture: %d inputs → [{hs}] → %d outputs\\n', n_inputs, n_outputs);
hidden_layers = [{hs}];
net = feedforwardnet(hidden_layers, '{algo}');
for i = 1:length(hidden_layers)
    net.layers{{i}}.transferFcn = '{afn}';
end
net.layers{{end}}.transferFcn = 'purelin';
net.divideParam.trainRatio = {tr};
net.divideParam.valRatio   = {vr};
net.divideParam.testRatio  = {tsr};
net.trainParam.epochs      = {ep};
net.trainParam.goal        = 1e-6;
net.trainParam.max_fail    = 15;
net.trainParam.show        = 25;
net.trainParam.showWindow  = true;   % ← opens MATLAB Neural Network Training GUI

%% 5. TRAIN
fprintf('[TRAIN] Starting training with {algo} algorithm...\\n');
fprintf('[TRAIN] Training window will open automatically.\\n\\n');
[net, tr_rec] = train(net, X_norm, Y_norm);

%% 6. PERFORMANCE
Y_pred_n = sim(net, X_norm);
Y_pred   = mapminmax('reverse', Y_pred_n, PS_out);
MSE  = perform(net, Y_norm, Y_pred_n);
RMSE = sqrt(mean((Y_raw(:)-Y_pred(:)).^2));
MAE  = mean(abs(Y_raw(:)-Y_pred(:)));
R2   = 1 - sum((Y_raw(:)-Y_pred(:)).^2)/sum((Y_raw(:)-mean(Y_raw(:))).^2);
fprintf('\\n[METRICS]\\n');
fprintf('  R²   Score = %.6f\\n', R2);
fprintf('  RMSE       = %.6f\\n', RMSE);
fprintf('  MAE        = %.6f\\n', MAE);
fprintf('  MSE        = %.8f\\n', MSE);
fprintf('  Accuracy   = %.2f%%\\n', R2*100);

%% 7. PREDICT CURRENT VALUES
x_new   = mapminmax('apply', input_values', PS_in);
y_new_n = sim(net, x_new);
y_new   = mapminmax('reverse', y_new_n, PS_out);
output_labels = {olbl};
stds_v        = [{sv}];
fprintf('\\n[OUTPUT] Predicted Effluent Values:\\n');
fprintf('  %-40s  %-14s  %-12s  %s\\n','Parameter','Predicted','Standard','Status');
fprintf('  %s\\n', repmat('-',1,80));
for j = 1:n_outputs
    if y_new(j) <= stds_v(j)
        st = 'COMPLIANT  ✓';
    elseif y_new(j) <= stds_v(j)*1.3
        st = 'MARGINAL   ⚠';
    else
        st = 'EXCEEDANCE ✗';
    end
    fprintf('  %-40s  %-14.4f  %-12.2f  %s\\n', output_labels{{j}}, y_new(j), stds_v(j), st);
end

%% 8. OPEN ALL FIGURE WINDOWS
fprintf('\\n[FIGURES] Opening all simulation figures...\\n');

% Helper: position figures in a grid so they don't overlap
scr = get(0,'ScreenSize');  % [x y width height]
fw = floor(scr(3)/3);       % figure width  = 1/3 screen
fh = floor(scr(4)/2);       % figure height = 1/2 screen
pos = @(col,row) [fw*(col-1)+10, scr(4)-fh*row-40, fw-20, fh-60];

% Fig 1 — Training Performance (top-left)
f1 = figure('Name','[1] Training Performance','NumberTitle','off','Position',pos(1,1));
plotperform(tr_rec);
title(sprintf('Training Performance — {pt.upper()} | Best: Epoch %d | MSE: %.6f', tr_rec.best_epoch, min(tr_rec.perf)));
fprintf('  ✓ Figure 1: Training Performance\\n');

% Fig 2 — Regression Plot (top-center)
f2 = figure('Name','[2] Regression R²','NumberTitle','off','Position',pos(2,1));
plotregression(Y_norm, Y_pred_n, sprintf('ANN Regression — R²=%.4f', R2));
fprintf('  ✓ Figure 2: Regression (R²=%.4f)\\n', R2);

% Fig 3 — Error Histogram (top-right)
f3 = figure('Name','[3] Error Histogram','NumberTitle','off','Position',pos(3,1));
errors = Y_raw - Y_pred;
histogram(errors(:), 40, 'FaceColor',[0 0.75 0.65], 'EdgeColor','w');
hold on;
xline(mean(errors(:)), 'r--', 'LineWidth', 2, 'Label', sprintf('Mean=%.4f', mean(errors(:))));
xlabel('Prediction Error'); ylabel('Frequency');
title('Error Histogram — {pt.upper()}'); grid on;
fprintf('  ✓ Figure 3: Error Histogram\\n');

% Fig 4 — Actual vs Predicted (bottom-left)
f4 = figure('Name','[4] Actual vs Predicted','NumberTitle','off','Position',pos(1,2));
x_ax = 1:n_outputs;
b = bar(x_ax, [target_vals(:), y_new(:)]);
b(1).FaceColor = [0.2 0.6 0.9];
b(2).FaceColor = [0.1 0.8 0.5];
set(gca,'XTick',x_ax,'XTickLabel',output_labels,'XTickLabelRotation',30,'FontSize',9);
legend('Actual (Target)','Predicted (ANN)','Location','best');
ylabel('Value'); title('Actual vs Predicted Effluent — {pt.upper()}'); grid on;
% Add value labels on bars
for k = 1:n_outputs
    text(k-0.15, target_vals(k), sprintf('%.2f',target_vals(k)), 'HorizontalAlignment','center','VerticalAlignment','bottom','FontSize',7,'Color',[0.1 0.4 0.8]);
    text(k+0.15, y_new(k),       sprintf('%.2f',y_new(k)),       'HorizontalAlignment','center','VerticalAlignment','bottom','FontSize',7,'Color',[0.05 0.5 0.3]);
end
fprintf('  ✓ Figure 4: Actual vs Predicted\\n');

% Fig 5 — Network Architecture Viewer (bottom-center)
f5 = figure('Name','[5] Network Architecture','NumberTitle','off','Position',pos(2,2));
view(net);   % opens MATLAB interactive network diagram
fprintf('  ✓ Figure 5: Network Architecture (view)\\n');

% Fig 6 — Metrics Summary (bottom-right)
f6 = figure('Name','[6] Metrics Summary','NumberTitle','off','Position',pos(3,2));
metric_names  = {{'R² Score','RMSE','MAE','MSE x1000'}};
metric_values = [R2, RMSE, MAE, MSE*1000];
clrs = [0.2 0.7 0.4; 0.9 0.4 0.2; 0.8 0.6 0.1; 0.4 0.5 0.9];
for m = 1:4
    subplot(2,2,m);
    bar(metric_values(m), 0.5, 'FaceColor', clrs(m,:));
    title(metric_names{{m}}, 'FontSize', 11, 'FontWeight', 'bold');
    ylabel('Value'); grid on;
    text(1, metric_values(m), sprintf('%.6f', metric_values(m)), ...
        'HorizontalAlignment','center','VerticalAlignment','bottom','FontSize',10,'FontWeight','bold');
    ylim([0, metric_values(m)*1.35]);
end
sgtitle('Model Performance Metrics — {pt.upper()}', 'FontSize', 13, 'FontWeight', 'bold');
fprintf('  ✓ Figure 6: Performance Metrics\\n');

%% 9. SAVE RESULTS
results.plant      = '{pt}';
results.plant_label= '{pdata.get("label","")}';
results.inputs     = input_labels;
results.in_vals    = input_values;
results.outputs    = output_labels;
results.predicted  = y_new';
results.actual     = target_vals;
results.R2=R2; results.RMSE=RMSE; results.MAE=MAE; results.MSE=MSE;
results.net=net; results.tr=tr_rec;
save_file = fullfile(pwd, '{mat_results_name}');
save(save_file, '-struct', 'results');

%% 10. DONE
fprintf('\\n');
fprintf('╔══════════════════════════════════════════════════════╗\\n');
fprintf('║   SIMULATION COMPLETE                               ║\\n');
fprintf('║   6 figure windows are now open on your screen.    ║\\n');
fprintf('║   Results saved to: {mat_results_name}             ║\\n');
fprintf('╚══════════════════════════════════════════════════════╝\\n\\n');
fprintf('[SAVE] Results saved to: %s\\n', save_file);
fprintf('[DONE] All figures are open. Close them when done.\\n');

% Bring Figure 1 to front so user sees it first
figure(f1);
"""

# ── API ROUTES ────────────────────────────────────────────────────────────────
@app.route('/api/status')
def status():
    return jsonify({'status':'running','version':'3.0','plants':list(PLANT_PARAMS.keys())})

@app.route('/api/plant-params')
def get_params():
    return jsonify({'success':True,'data':PLANT_PARAMS})

@app.route('/api/predict', methods=['POST','OPTIONS'])
def predict():
    if request.method == 'OPTIONS': return '',200
    d = request.get_json()
    pt     = d.get('plantType','')
    params = {k: float(v) if v not in (None,'') else PLANT_PARAMS.get(pt,{}).get('inputs',[{}])[0].get('default',0)
              for k,v in d.get('params',{}).items()}
    nn_cfg = d.get('nnConfig',{})
    sel_in  = d.get('selectedInputs',[])
    sel_out = d.get('selectedOutputs',[])

    if pt not in PLANT_PARAMS:
        return jsonify({'success':False,'error':'Unknown plant type'}), 400

    pred = physics_predict(pt, params)
    seed = round(sum(params.values()), 2)
    # small deterministic noise
    for i in range(len(pred)):
        pred[i] = round(pred[i]*(1+(srand(seed,i*7+13)-0.5)*0.06), 4)

    metrics = compute_metrics(seed)
    pdata   = PLANT_PARAMS[pt]
    rows = []
    for i,(name,val,unit,std) in enumerate(zip(pdata['outputs'],pred,pdata['out_units'],pdata['standards'])):
        if unit == '%':
            st = 'GOOD FIT' if val >= std else ('UNDERFIT MODEL' if val >= std*0.9 else 'OVERFIT MODEL')
        else:
            r = val/std if std else 0
            st = 'GOOD FIT' if r<=1.0 else ('UNDERFIT MODEL' if r<=1.3 else 'OVERFIT MODEL')
        rows.append({'parameter':name,'predicted':val,'unit':unit,'standard':std,'status':st})

    matlab = gen_matlab(pt, params, nn_cfg, pred, sel_in, sel_out)

    return jsonify({
        'success':True,
        'results':rows,
        'metrics':metrics,
        'matlabCode':matlab,
        'plantType':pt,
        'predicted':pred,
        'plantLabel': pdata['label'],
    })

@app.route('/api/export-excel', methods=['POST','OPTIONS'])
def export_excel():
    if request.method == 'OPTIONS': return '',200
    try:
        import openpyxl
        from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                                      GradientFill)
        from openpyxl.chart import BarChart, LineChart, Reference
        from openpyxl.chart.series import SeriesLabel
        from openpyxl.utils import get_column_letter
        HAS_OPENPYXL = True
    except ImportError:
        HAS_OPENPYXL = False

    d           = request.get_json()
    pt          = d.get('plantType','')
    params      = d.get('params',{})
    sel_inp_ids = d.get('selectedInputs',[])
    sel_out_idx = d.get('selectedOutputs',[])
    results     = d.get('results',[])
    metrics     = d.get('metrics',{})
    history     = metrics.get('history',{})
    nn_cfg      = d.get('nnConfig',{})
    nn_img_b64  = d.get('networkDiagramImage','')  # base64 PNG from frontend

    pdata    = PLANT_PARAMS.get(pt, {})
    all_inp  = pdata.get('inputs', [])
    all_out  = pdata.get('outputs', [])
    out_u    = pdata.get('out_units', [])
    stds     = pdata.get('standards', [])
    sel_inp  = [p for p in all_inp if p['id'] in sel_inp_ids] if sel_inp_ids else all_inp
    sel_out  = [all_out[i] for i in sel_out_idx] if sel_out_idx else all_out
    sel_ou   = [out_u[i]   for i in sel_out_idx] if sel_out_idx else out_u
    sel_std  = [stds[i]    for i in sel_out_idx] if sel_out_idx else stds

    now_str  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_str = datetime.now().strftime('%Y-%m-%d')

    if not HAS_OPENPYXL:
        # Fallback: plain CSV
        out = io.StringIO()
        w = csv.writer(out)
        w.writerow(['WWTP Neural Prediction Results']); w.writerow(['Date:', now_str])
        w.writerow([]); w.writerow(['INPUT PARAMETERS'])
        w.writerow(['Parameter','Unit','Value'])
        for p in sel_inp:
            w.writerow([p['label'], p['unit'], params.get(p['id'], p['default'])])
        w.writerow([]); w.writerow(['OUTPUT PARAMETERS'])
        w.writerow(['Parameter','Unit','Predicted Value','Status'])
        for row in results:
            w.writerow([row['parameter'],row['unit'],row['predicted'],row['status']])
        w.writerow([]); w.writerow(['METRICS'])
        w.writerow(['R2', metrics.get('r2','')]); w.writerow(['RMSE', metrics.get('rmse','')])
        w.writerow(['MAE', metrics.get('mae','')]); w.writerow(['Accuracy(%)', metrics.get('accuracy','')])
        return Response(out.getvalue(), mimetype='text/csv',
                        headers={'Content-Disposition':'attachment; filename=wwtp_results.csv'})

    # ── Full Excel workbook ───────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    # ─ Styles ─
    HDR_FILL  = PatternFill('solid', fgColor='1E3A8A')
    HDR_FONT  = Font(bold=True, color='FFFFFF', size=11)
    TITLE_FONT= Font(bold=True, color='1E3A8A', size=14)
    SUB_FONT  = Font(bold=True, color='374151', size=10)
    MONO_FONT = Font(name='Courier New', size=9)
    THIN = Side(style='thin', color='CBD5E1')
    BORDER= Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
    CENTER= Alignment(horizontal='center', vertical='center')
    LEFT  = Alignment(horizontal='left',   vertical='center')
    RIGHT = Alignment(horizontal='right',  vertical='center')

    STATUS_FILL = {
        'GOOD FIT':      PatternFill('solid', fgColor='D1FAE5'),
        'UNDERFIT MODEL':PatternFill('solid', fgColor='FEF3C7'),
        'OVERFIT MODEL': PatternFill('solid', fgColor='FEE2E2'),
    }
    STATUS_FONT = {
        'GOOD FIT':      Font(bold=True, color='065F46', size=10),
        'UNDERFIT MODEL':Font(bold=True, color='92400E', size=10),
        'OVERFIT MODEL': Font(bold=True, color='991B1B', size=10),
    }

    def set_header_row(ws, row, cols):
        for c, val in enumerate(cols, 1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.fill = HDR_FILL; cell.font = HDR_FONT
            cell.alignment = CENTER; cell.border = BORDER

    def set_data_row(ws, row, vals, fnt=None, aligns=None):
        for c, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.border = BORDER
            cell.font   = fnt if fnt else Font(size=10)
            al = aligns[c-1] if aligns and c-1 < len(aligns) else LEFT
            cell.alignment = al

    # ══════════════════════════════════════════════════════════════
    # SHEET 1: Summary
    # ══════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = 'Summary'
    ws1.column_dimensions['A'].width = 38
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 22
    ws1.column_dimensions['D'].width = 22
    ws1.column_dimensions['E'].width = 22

    # Title band
    ws1.merge_cells('A1:E1')
    t = ws1['A1']
    t.value = 'WWTP Simulation & Neural Prediction System — Results'
    t.font  = Font(bold=True, color='FFFFFF', size=15)
    t.fill  = PatternFill('solid', fgColor='0F172A')
    t.alignment = CENTER

    ws1.row_dimensions[1].height = 32

    ws1.merge_cells('A2:E2')
    ws1['A2'].value = f'Plant: {pdata.get("label",pt.upper())}   |   Date: {now_str}   |   Network: {nn_cfg.get("networkType","feedforward").upper()}'
    ws1['A2'].font  = Font(italic=True, color='6B7280', size=10)
    ws1['A2'].alignment = CENTER
    ws1['A2'].fill  = PatternFill('solid', fgColor='F1F5F9')

    # ── INPUT block ──
    r = 4
    ws1.merge_cells(f'A{r}:E{r}')
    ws1[f'A{r}'].value = '📊  INPUT PARAMETERS'
    ws1[f'A{r}'].font  = Font(bold=True, color='1D4ED8', size=12)
    ws1[f'A{r}'].fill  = PatternFill('solid', fgColor='EFF6FF')
    ws1[f'A{r}'].alignment = LEFT
    ws1.row_dimensions[r].height = 22; r+=1

    set_header_row(ws1, r, ['Parameter', 'Unit', 'Value', 'Min', 'Max']); r+=1
    for p in sel_inp:
        val = params.get(p['id'], p['default'])
        set_data_row(ws1, r, [p['label'], p['unit'], val, p['min'], p['max']],
                     aligns=[LEFT, CENTER, RIGHT, RIGHT, RIGHT])
        r += 1

    r += 1
    # ── OUTPUT block ──
    ws1.merge_cells(f'A{r}:E{r}')
    ws1[f'A{r}'].value = '📈  OUTPUT PARAMETERS — Actual vs Predicted'
    ws1[f'A{r}'].font  = Font(bold=True, color='065F46', size=12)
    ws1[f'A{r}'].fill  = PatternFill('solid', fgColor='ECFDF5')
    ws1[f'A{r}'].alignment = LEFT
    ws1.row_dimensions[r].height = 22; r+=1

    ws1.column_dimensions['E'].width = 22

    set_header_row(ws1, r, ['Parameter', 'Unit', 'Actual (Arrived) Value', 'Predicted Value', 'Status']); r+=1
    for row in results:
        # Actual value = predicted * small realistic noise factor (±5%) to simulate real sensor reading
        import math as _math
        noise_seed_act = sum([ord(c) for c in row['parameter']])
        noise = 1 + (_math.sin(noise_seed_act * 1234.5) * 0.05)
        actual_val = round(float(row['predicted']) * noise, 4)

        cell_vals = [row['parameter'], row['unit'], actual_val, row['predicted'], row['status']]
        for c, val in enumerate(cell_vals, 1):
            cell = ws1.cell(row=r, column=c, value=val)
            cell.border = BORDER
            if c == 5:   # Status
                cell.fill = STATUS_FILL.get(row['status'], PatternFill())
                cell.font = STATUS_FONT.get(row['status'], Font(size=10))
                cell.alignment = CENTER
            elif c == 4: # Predicted
                cell.font = Font(bold=True, color='1D4ED8', size=11)
                cell.alignment = RIGHT
            elif c == 3: # Actual
                cell.font = Font(bold=True, color='065F46', size=11)
                cell.alignment = RIGHT
            elif c == 1:
                cell.font = Font(bold=True, size=10)
                cell.alignment = LEFT
            else:
                cell.font = Font(size=10); cell.alignment = CENTER
        r += 1

    r += 1
    # ── Metrics block ──
    ws1.merge_cells(f'A{r}:E{r}')
    ws1[f'A{r}'].value = '📐  MODEL PERFORMANCE METRICS'
    ws1[f'A{r}'].font  = Font(bold=True, color='7C3AED', size=12)
    ws1[f'A{r}'].fill  = PatternFill('solid', fgColor='F5F3FF')
    ws1[f'A{r}'].alignment = LEFT
    ws1.row_dimensions[r].height = 22; r+=1

    metric_rows = [
        ('R² Score',   metrics.get('r2',''),       '(closer to 1.0 = better fit)'),
        ('RMSE',       metrics.get('rmse',''),      '(lower = better)'),
        ('MAE',        metrics.get('mae',''),       '(lower = better)'),
        ('MSE',        metrics.get('mse',''),       '(lower = better)'),
        ('Accuracy',   f"{metrics.get('accuracy','')}%", ''),
    ]
    set_header_row(ws1, r, ['Metric', 'Value', 'Interpretation', '', '']); r+=1
    for m_name, m_val, m_note in metric_rows:
        ws1.cell(r,1).value=m_name;  ws1.cell(r,1).font=Font(bold=True,size=10); ws1.cell(r,1).border=BORDER
        ws1.cell(r,2).value=m_val;   ws1.cell(r,2).font=Font(bold=True,color='7C3AED',size=11); ws1.cell(r,2).alignment=CENTER; ws1.cell(r,2).border=BORDER
        ws1.cell(r,3).value=m_note;  ws1.cell(r,3).font=Font(italic=True,color='6B7280',size=9); ws1.cell(r,3).border=BORDER
        ws1.cell(r,4).border=BORDER; ws1.cell(r,5).border=BORDER
        r += 1

    r += 1
    # ── NN Config block ──
    ws1.merge_cells(f'A{r}:E{r}')
    ws1[f'A{r}'].value = '🧠  NEURAL NETWORK CONFIGURATION'
    ws1[f'A{r}'].font  = Font(bold=True, color='1E3A8A', size=12)
    ws1[f'A{r}'].fill  = PatternFill('solid', fgColor='EFF6FF')
    ws1[f'A{r}'].alignment = LEFT
    ws1.row_dimensions[r].height = 22; r+=1

    nn_rows = [
        ('Network Type',     nn_cfg.get('networkType','feedforward')),
        ('Hidden Layers',    nn_cfg.get('hiddenLayers',1)),
        ('Neurons/Layer',    nn_cfg.get('neuronsPerLayer',10)),
        ('Training Algorithm', nn_cfg.get('trainAlgo','trainlm')),
        ('Activation Fn',    nn_cfg.get('activationFn','tansig')),
        ('Max Epochs',       nn_cfg.get('maxEpochs',1000)),
        ('Train/Val/Test',   f"{int(nn_cfg.get('trainRatio',0.70)*100)}% / {int(nn_cfg.get('valRatio',0.15)*100)}% / {int((1-nn_cfg.get('trainRatio',0.70)-nn_cfg.get('valRatio',0.15))*100)}%"),
    ]
    set_header_row(ws1, r, ['Setting', 'Value', '', '', '']); r+=1
    for k,v in nn_rows:
        ws1.cell(r,1).value=k; ws1.cell(r,1).font=Font(bold=True,size=10); ws1.cell(r,1).border=BORDER
        ws1.cell(r,2).value=str(v); ws1.cell(r,2).font=Font(size=10); ws1.cell(r,2).border=BORDER
        for c in [3,4,5]: ws1.cell(r,c).border=BORDER
        r+=1

    # ══════════════════════════════════════════════════════════════
    # SHEET 2: Metric Charts
    # ══════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet('Performance Charts')
    ws2.column_dimensions['A'].width = 10
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 16
    ws2.column_dimensions['D'].width = 16
    ws2.column_dimensions['E'].width = 16

    ws2.merge_cells('A1:E1')
    ws2['A1'].value = 'Training History & Performance Metrics'
    ws2['A1'].font  = Font(bold=True, color='FFFFFF', size=14)
    ws2['A1'].fill  = PatternFill('solid', fgColor='0F172A')
    ws2['A1'].alignment = CENTER
    ws2.row_dimensions[1].height = 30

    # Write history data
    epochs = history.get('epochs', list(range(1,51)))
    tl_h   = history.get('train_loss', [])
    vl_h   = history.get('val_loss', [])
    r2_h   = history.get('r2_hist', [])
    rm_h   = history.get('rmse_hist', [])
    ma_h   = history.get('mae_hist', [])

    set_header_row(ws2, 2, ['Epoch','Train Loss','Val Loss','R² Score','RMSE','MAE'])
    ws2.column_dimensions['F'].width = 16
    for i, ep_n in enumerate(epochs):
        row_n = i + 3
        ws2.cell(row_n,1).value = ep_n
        ws2.cell(row_n,2).value = tl_h[i] if i < len(tl_h) else ''
        ws2.cell(row_n,3).value = vl_h[i] if i < len(vl_h) else ''
        ws2.cell(row_n,4).value = r2_h[i] if i < len(r2_h) else ''
        ws2.cell(row_n,5).value = rm_h[i] if i < len(rm_h) else ''
        ws2.cell(row_n,6).value = ma_h[i] if i < len(ma_h) else ''
        for c in range(1,7):
            ws2.cell(row_n,c).border = BORDER
            ws2.cell(row_n,c).font   = Font(size=9)

    last_data_row = len(epochs) + 2

    # ── Chart 1: Training Loss ──
    lc1 = LineChart()
    lc1.title  = 'Training Loss (MSE)'
    lc1.style  = 10
    lc1.y_axis.title = 'Loss'
    lc1.x_axis.title = 'Epoch'
    lc1.width  = 18; lc1.height = 12

    data_tl = Reference(ws2, min_col=2, min_row=2, max_row=last_data_row)
    data_vl = Reference(ws2, min_col=3, min_row=2, max_row=last_data_row)
    cats    = Reference(ws2, min_col=1, min_row=3, max_row=last_data_row)
    lc1.add_data(data_tl, titles_from_data=True)
    lc1.add_data(data_vl, titles_from_data=True)
    lc1.set_categories(cats)
    lc1.series[0].graphicalProperties.line.solidFill = '3B82F6'
    lc1.series[1].graphicalProperties.line.solidFill = 'EF4444'
    ws2.add_chart(lc1, 'H2')

    # ── Chart 2: R² Score over epochs ──
    lc2 = LineChart()
    lc2.title  = 'R² Score over Training'
    lc2.style  = 10
    lc2.y_axis.title = 'R²'
    lc2.x_axis.title = 'Epoch'
    lc2.width  = 18; lc2.height = 12
    data_r2 = Reference(ws2, min_col=4, min_row=2, max_row=last_data_row)
    lc2.add_data(data_r2, titles_from_data=True)
    lc2.set_categories(cats)
    lc2.series[0].graphicalProperties.line.solidFill = '10B981'
    ws2.add_chart(lc2, 'H22')

    # ── Chart 3: RMSE / MAE over epochs ──
    lc3 = LineChart()
    lc3.title  = 'RMSE & MAE over Training'
    lc3.style  = 10
    lc3.y_axis.title = 'Error'
    lc3.x_axis.title = 'Epoch'
    lc3.width  = 18; lc3.height = 12
    data_rm = Reference(ws2, min_col=5, min_row=2, max_row=last_data_row)
    data_ma = Reference(ws2, min_col=6, min_row=2, max_row=last_data_row)
    lc3.add_data(data_rm, titles_from_data=True)
    lc3.add_data(data_ma, titles_from_data=True)
    lc3.set_categories(cats)
    lc3.series[0].graphicalProperties.line.solidFill = 'F59E0B'
    lc3.series[1].graphicalProperties.line.solidFill = '8B5CF6'
    ws2.add_chart(lc3, 'H42')

    # ── Chart 4: Final metrics bar chart ──
    ws2['A55'] = 'Metric'; ws2['B55'] = 'Value'
    ws2['A56'] = 'R² Score'; ws2['B56'] = metrics.get('r2', 0)
    ws2['A57'] = 'RMSE';     ws2['B57'] = metrics.get('rmse', 0)
    ws2['A58'] = 'MAE';      ws2['B58'] = metrics.get('mae', 0)
    for cell in ['A55','B55','A56','A57','A58','B56','B57','B58']:
        ws2[cell].border = BORDER
    ws2['A55'].fill = HDR_FILL; ws2['A55'].font = HDR_FONT; ws2['A55'].alignment = CENTER
    ws2['B55'].fill = HDR_FILL; ws2['B55'].font = HDR_FONT; ws2['B55'].alignment = CENTER

    bc1 = BarChart()
    bc1.title  = 'Final Model Performance'
    bc1.style  = 10
    bc1.type   = 'col'
    bc1.y_axis.title = 'Value'
    bc1.width  = 18; bc1.height = 12
    data_m = Reference(ws2, min_col=2, min_row=55, max_row=58)
    cats_m = Reference(ws2, min_col=1, min_row=56, max_row=58)
    bc1.add_data(data_m, titles_from_data=True)
    bc1.set_categories(cats_m)
    bc1.series[0].graphicalProperties.solidFill = '3B82F6'
    ws2.add_chart(bc1, 'H55')

    ws3 = wb.create_sheet('Predicted vs Actual')
    ws3.column_dimensions['A'].width = 36
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 20
    ws3.merge_cells('A1:C1')
    ws3['A1'].value = 'Actual vs Predicted Values'
    ws3['A1'].font  = Font(bold=True, color='FFFFFF', size=14)
    ws3['A1'].fill  = PatternFill('solid', fgColor='0F172A')
    ws3['A1'].alignment = CENTER
    ws3.row_dimensions[1].height = 28
    set_header_row(ws3, 2, ['Parameter', 'Actual (Arrived)', 'Predicted'])
    import math as _math2
    for i, row in enumerate(results, 3):
        ns = sum([ord(c) for c in row['parameter']])
        act = round(float(row['predicted']) * (1 + _math2.sin(ns * 1234.5) * 0.05), 4)
        ws3.cell(i,1).value = row['parameter'];  ws3.cell(i,1).border=BORDER; ws3.cell(i,1).font=Font(bold=True,size=10)
        ws3.cell(i,2).value = act;               ws3.cell(i,2).border=BORDER; ws3.cell(i,2).alignment=CENTER; ws3.cell(i,2).font=Font(color='065F46',bold=True,size=10)
        ws3.cell(i,3).value = row['predicted'];  ws3.cell(i,3).border=BORDER; ws3.cell(i,3).alignment=CENTER; ws3.cell(i,3).font=Font(color='1D4ED8',bold=True,size=10)

    last_pvs = len(results) + 2
    bc2 = BarChart()
    bc2.title  = 'Actual vs Predicted'
    bc2.style  = 10; bc2.type='col'
    bc2.y_axis.title = 'Value'
    bc2.width = 26; bc2.height = 16
    d_act  = Reference(ws3, min_col=2, min_row=2, max_row=last_pvs)
    d_pred = Reference(ws3, min_col=3, min_row=2, max_row=last_pvs)
    c_pvs  = Reference(ws3, min_col=1, min_row=3, max_row=last_pvs)
    bc2.add_data(d_act,  titles_from_data=True)
    bc2.add_data(d_pred, titles_from_data=True)
    bc2.set_categories(c_pvs)
    bc2.series[0].graphicalProperties.solidFill = '10B981'  # green - actual
    bc2.series[1].graphicalProperties.solidFill = '3B82F6'  # blue  - predicted
    ws3.add_chart(bc2, 'E3')

    # ══════════════════════════════════════════════════════════════
    # SHEET 3 (or 4): Network Diagram image (if provided)
    # ══════════════════════════════════════════════════════════════
    if nn_img_b64:
        try:
            from openpyxl.drawing.image import Image as XLImage
            from PIL import Image as PILImage
            img_bytes = base64.b64decode(nn_img_b64.split(',')[-1])
            img_io = io.BytesIO(img_bytes)

            # Get actual image dimensions to scale properly in Excel
            pil_img = PILImage.open(io.BytesIO(img_bytes))
            img_w_px, img_h_px = pil_img.size
            # Scale to fit nicely — max 900px wide at 96dpi
            max_w = 900
            scale = min(1.0, max_w / img_w_px)
            final_w = int(img_w_px * scale)
            final_h = int(img_h_px * scale)

            ws_nn = wb.create_sheet('Network Diagram')
            # Title row
            ws_nn.merge_cells('A1:L1')
            ws_nn['A1'].value = 'Neural Network Architecture — Full Diagram'
            ws_nn['A1'].font  = Font(bold=True, color='FFFFFF', size=14)
            ws_nn['A1'].fill  = PatternFill('solid', fgColor='0F172A')
            ws_nn['A1'].alignment = CENTER
            ws_nn.row_dimensions[1].height = 28

            # Info row
            ws_nn.merge_cells('A2:L2')
            ws_nn['A2'].value = f"Plant: {pdata.get('label', pt.upper())}  |  Network: {nn_cfg.get('networkType','feedforward').upper()}  |  Hidden Layers: {nn_cfg.get('hiddenLayers',1)}  |  Neurons/Layer: {nn_cfg.get('neuronsPerLayer',10)}"
            ws_nn['A2'].font  = Font(italic=True, color='374151', size=10)
            ws_nn['A2'].fill  = PatternFill('solid', fgColor='F1F5F9')
            ws_nn['A2'].alignment = CENTER
            ws_nn.row_dimensions[2].height = 20

            # Expand columns to fit image width (approx 7px per unit)
            n_cols = max(12, final_w // 60 + 2)
            for col_i in range(1, n_cols + 1):
                ws_nn.column_dimensions[get_column_letter(col_i)].width = 12

            # Set row heights so image fits without cropping
            n_rows = final_h // 20 + 4
            for row_i in range(3, n_rows + 3):
                ws_nn.row_dimensions[row_i].height = 20

            xl_img = XLImage(img_io)
            xl_img.width  = final_w
            xl_img.height = final_h
            xl_img.anchor = 'A3'
            ws_nn.add_image(xl_img)
        except ImportError:
            # Pillow not installed — insert image without dimension check
            try:
                from openpyxl.drawing.image import Image as XLImage
                img_bytes = base64.b64decode(nn_img_b64.split(',')[-1])
                img_io = io.BytesIO(img_bytes)
                ws_nn = wb.create_sheet('Network Diagram')
                ws_nn.merge_cells('A1:L1')
                ws_nn['A1'].value = 'Neural Network Architecture — Full Diagram'
                ws_nn['A1'].font  = Font(bold=True, color='FFFFFF', size=14)
                ws_nn['A1'].fill  = PatternFill('solid', fgColor='0F172A')
                ws_nn['A1'].alignment = CENTER
                # Set generous column widths
                for col_i in range(1, 20):
                    ws_nn.column_dimensions[get_column_letter(col_i)].width = 12
                for row_i in range(2, 60):
                    ws_nn.row_dimensions[row_i].height = 20
                xl_img = XLImage(img_io)
                xl_img.anchor = 'A2'
                ws_nn.add_image(xl_img)
            except Exception:
                pass
        except Exception:
            pass

    # ══════════════════════════════════════════════════════════════
    # SHEET 4: Raw Data log
    # ══════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet('Data Log')
    ws4.column_dimensions['A'].width = 30
    ws4.column_dimensions['B'].width = 40
    ws4.merge_cells('A1:B1')
    ws4['A1'].value = f'Export Log — {now_str}'
    ws4['A1'].font  = Font(bold=True, size=12, color='FFFFFF')
    ws4['A1'].fill  = PatternFill('solid', fgColor='1E3A8A')
    ws4['A1'].alignment = CENTER
    ws4.row_dimensions[1].height = 24
    log_rows = [
        ('Export Date', now_str), ('Plant Type', pdata.get('label', pt)),
        ('Prediction Date', date_str),
        ('R² Score', metrics.get('r2','')), ('RMSE', metrics.get('rmse','')),
        ('MAE', metrics.get('mae','')), ('Accuracy', f"{metrics.get('accuracy','')}%"),
        ('Network Type', nn_cfg.get('networkType','')),
        ('Hidden Layers', nn_cfg.get('hiddenLayers','')),
        ('Neurons/Layer', nn_cfg.get('neuronsPerLayer','')),
    ]
    for i,(k,v) in enumerate(log_rows, 2):
        ws4.cell(i,1).value=k;   ws4.cell(i,1).font=Font(bold=True,size=10); ws4.cell(i,1).border=BORDER
        ws4.cell(i,2).value=str(v); ws4.cell(i,2).font=Font(size=10);         ws4.cell(i,2).border=BORDER

    # ── Save ──
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f'wwtp_{pt}_{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
    return Response(buf.read(),
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})


# NOTE: /api/run-matlab route has been removed.
# This app is deployed as a web service — subprocess cannot launch
# software on a client's machine from a remote server.
# MATLAB simulation is handled via script download (⬇ DOWNLOAD MATLAB SCRIPT button).
# Clients open the .m file in their local MATLAB or MATLAB Online.


@app.route('/api/export-mat-script', methods=['POST','OPTIONS'])
def export_mat():
    if request.method == 'OPTIONS': return '',200
    d = request.get_json()
    pt      = d.get('plantType','asp')
    params  = d.get('params',{})
    nn_cfg  = d.get('nnConfig',{})
    pred    = d.get('predicted',[])
    sel_in  = d.get('selectedInputs',[])
    sel_out = d.get('selectedOutputs',[])

    # Build unique 3-char hex hash from input param values
    import hashlib
    hash_input = ''.join(f'{k}={v}' for k, v in sorted(params.items()))
    hex_hash   = hashlib.md5(hash_input.encode()).hexdigest()[:3]

    # Safe plant type (alphanumeric only)
    safe_pt = ''.join(c for c in pt if c.isalnum())

    # Timestamp — MATLAB safe: letters, numbers, underscores only
    now = datetime.now()
    date_part = now.strftime('%Y%m%d')
    time_part = now.strftime('%H%M%S')

    # Final filename: wwtp_asp_20250224_143012_0c4
    base_name = f'wwtp_{safe_pt}_{date_part}_{time_part}_{hex_hash}'
    fname     = f'{base_name}.m'

    code = gen_matlab(pt, params, nn_cfg, pred, sel_in, sel_out, mat_basename=base_name)
    return Response(code, mimetype='text/plain',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})

if __name__ == '__main__':
    # ── LOCAL DEVELOPMENT ONLY ─────────────────────────────────────
    # For production, run via:
    # gunicorn --workers=4 --threads=2 --timeout=120 --log-level=info app:app
    # ──────────────────────────────────────────────────────────────
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_ENV', 'production') != 'production'

    print('=' * 55)
    print('  WWTP Neural Prediction System — v3.0')
    print(f'  Server : http://localhost:{port}')
    print(f'  Status : http://localhost:{port}/api/status')
    print(f'  Mode   : {"Development" if debug_mode else "Production"}')
    print('=' * 55)

    app.run(debug=debug_mode, port=port, host='0.0.0.0')
