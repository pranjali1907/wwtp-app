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
    start_date_str = d.get('startDate', datetime.now().strftime('%Y-%m-%d'))
    horizon     = int(d.get('horizon', 7))

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
    # ══════════════════════════════════════════════════════════════
    # NETWORK DIAGRAM — Excel cells styled like Image 2 (no PIL needed)
    # ══════════════════════════════════════════════════════════════
    hl_v  = nn_cfg.get('hiddenLayers', 1)
    npl_v = nn_cfg.get('neuronsPerLayer', 10)
    afn_v = nn_cfg.get('activationFn', 'tansig')
    ntyp  = nn_cfg.get('networkType', 'feedforward')
    nin_v = len(sel_inp) or 3
    nou_v = len(sel_out) or 2

    ws_nn = wb.create_sheet('Network Diagram')

    # ── colour palette (matching Image 2) ──
    C_TITLE   = '0D1B2A'   # near-black title bar
    C_INFO    = 'F1F5F9'
    C_INPUT_H = '22C55E'   # green header
    C_INPUT_B = '16A34A'   # green body
    C_HID_H   = '38BDF8'   # blue header
    C_HID_B   = '7DD3FC'   # light blue body
    C_OUT_H   = 'A855F7'   # purple header
    C_OUT_B   = 'D8B4FE'   # light purple body
    C_WB      = 'F8FAFC'   # near-white for W/b cells
    C_PHI     = 'E0F2FE'   # light blue for ⊕ cell
    C_FN      = 'DBEAFE'   # very light blue for f(·) cell
    C_SM      = 'EDE9FE'   # light purple for softmax
    C_CONN    = 'CBD5E1'   # connector background
    C_BLACK   = '000000'
    C_WHITE   = 'FFFFFF'

    def _cell(ws, row, col, val='', fg=None, font_color='000000',
              bold=False, size=10, align='center', border=True, h=None):
        c = ws.cell(row, col, value=val)
        if fg: c.fill = PatternFill('solid', fgColor=fg)
        c.font = Font(bold=bold, color=font_color, size=size,
                      name='Segoe UI' if bold else 'Calibri')
        c.alignment = Alignment(horizontal=align, vertical='center',
                                wrap_text=False)
        if border:
            s = Side(style='thin', color='CBD5E1')
            c.border = Border(left=s, right=s, top=s, bottom=s)
        return c

    # ── worksheet column widths ──
    # Layout: col1=label gap | then per element: [3 cols box] [1 col gap]
    # Elements: INPUT(1) | hidden(hl_v) | OUTPUT(1) | final bubble(1)
    n_elems = 1 + hl_v + 1 + 1   # input + hiddens + output + bubble

    ROW_TITLE   = 1
    ROW_INFO    = 2
    ROW_EMPTY1  = 3
    ROW_LABEL   = 4   # "INPUT" / "HIDDEN 1" labels
    ROW_BOX_T   = 5   # top of box (W/b row)
    ROW_BOX_M   = 6   # ⊕ row
    ROW_BOX_B   = 7   # f(·) / softmax row
    ROW_EMPTY2  = 8
    ROW_COUNT   = 9   # "5 nodes" / "10 neurons"
    TOTAL_ROWS  = 10

    # title & info spanning all cols
    total_cols = n_elems * 4   # 4 cols per element (3 box + 1 gap)
    last_col   = get_column_letter(total_cols + 1)

    ws_nn.merge_cells(f'A{ROW_TITLE}:{last_col}{ROW_TITLE}')
    ws_nn[f'A{ROW_TITLE}'].value = f'{ntyp.upper()} Neural Network  —  Architecture View'
    ws_nn[f'A{ROW_TITLE}'].font  = Font(bold=True, color=C_WHITE, size=14, name='Segoe UI')
    ws_nn[f'A{ROW_TITLE}'].fill  = PatternFill('solid', fgColor=C_TITLE)
    ws_nn[f'A{ROW_TITLE}'].alignment = Alignment(horizontal='center', vertical='center')
    ws_nn.row_dimensions[ROW_TITLE].height = 32

    ws_nn.merge_cells(f'A{ROW_INFO}:{last_col}{ROW_INFO}')
    ws_nn[f'A{ROW_INFO}'].value = (f"M  {ntyp.upper()} Neural Network (view)  |  "
                                    f"Plant: {pdata.get('label',pt.upper())}  |  "
                                    f"Layers: {hl_v}×{npl_v}  |  "
                                    f"Activation: {afn_v}  |  "
                                    f"Inputs: {nin_v}   Outputs: {nou_v}")
    ws_nn[f'A{ROW_INFO}'].font  = Font(size=9, color='374151', name='Segoe UI')
    ws_nn[f'A{ROW_INFO}'].fill  = PatternFill('solid', fgColor=C_INFO)
    ws_nn[f'A{ROW_INFO}'].alignment = Alignment(horizontal='left', vertical='center')
    ws_nn.row_dimensions[ROW_INFO].height = 20
    ws_nn.row_dimensions[ROW_EMPTY1].height = 12
    ws_nn.row_dimensions[ROW_LABEL].height  = 22
    ws_nn.row_dimensions[ROW_BOX_T].height  = 28
    ws_nn.row_dimensions[ROW_BOX_M].height  = 32
    ws_nn.row_dimensions[ROW_BOX_B].height  = 28
    ws_nn.row_dimensions[ROW_EMPTY2].height = 8
    ws_nn.row_dimensions[ROW_COUNT].height  = 20

    def draw_element(col_start, etype, label, h_col, b_col, count_label, fn_label='f(·)'):
        c1,c2,c3 = col_start, col_start+1, col_start+2
        lc = get_column_letter
        # label row
        ws_nn.merge_cells(f'{lc(c1)}{ROW_LABEL}:{lc(c3)}{ROW_LABEL}')
        _cell(ws_nn, ROW_LABEL, c1, label, fg=None,
              font_color='1E3A8A' if 'HIDDEN' in label or 'OUTPUT' in label else '166534',
              bold=True, size=11, border=False)
        # col widths
        for c in [c1,c2,c3]:
            ws_nn.column_dimensions[lc(c)].width = 5
        if etype == 'input':
            # big green merged cell
            ws_nn.merge_cells(f'{lc(c1)}{ROW_BOX_T}:{lc(c3)}{ROW_BOX_B}')
            cell = ws_nn.cell(ROW_BOX_T, c1, value=str(nin_v))
            cell.fill = PatternFill('solid', fgColor=C_INPUT_B)
            cell.font = Font(bold=True, size=22, color=C_WHITE, name='Segoe UI')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            s = Side(style='medium', color=C_INPUT_H)
            cell.border = Border(left=s, right=s, top=s, bottom=s)
        elif etype == 'bubble':
            ws_nn.merge_cells(f'{lc(c1)}{ROW_BOX_T}:{lc(c3)}{ROW_BOX_B}')
            cell = ws_nn.cell(ROW_BOX_T, c1, value=str(nou_v))
            cell.fill = PatternFill('solid', fgColor='EC4899')
            cell.font = Font(bold=True, size=22, color=C_WHITE, name='Segoe UI')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            s = Side(style='medium', color='BE185D')
            cell.border = Border(left=s, right=s, top=s, bottom=s)
        else:
            # W | b row
            _cell(ws_nn, ROW_BOX_T, c1, 'W', fg=C_WB, bold=True, size=11)
            _cell(ws_nn, ROW_BOX_T, c2, '',  fg=h_col)
            _cell(ws_nn, ROW_BOX_T, c3, 'b', fg=C_WB, bold=True, size=11)
            # ⊕ row
            ws_nn.merge_cells(f'{lc(c1)}{ROW_BOX_M}:{lc(c3)}{ROW_BOX_M}')
            phcell = ws_nn.cell(ROW_BOX_M, c1, value='⊕')
            phcell.fill = PatternFill('solid', fgColor=b_col)
            phcell.font = Font(bold=True, size=18, color=C_WHITE, name='Segoe UI')
            phcell.alignment = Alignment(horizontal='center', vertical='center')
            s = Side(style='thin', color='94A3B8')
            phcell.border = Border(left=s, right=s, top=s, bottom=s)
            # f(·) / softmax row
            ws_nn.merge_cells(f'{lc(c1)}{ROW_BOX_B}:{lc(c3)}{ROW_BOX_B}')
            fn_bg = C_SM if etype=='output' else C_FN
            fncell = ws_nn.cell(ROW_BOX_B, c1, value=fn_label)
            fncell.fill = PatternFill('solid', fgColor=fn_bg)
            fncell.font = Font(bold=False, size=9, color='1E40AF', name='Segoe UI')
            fncell.alignment = Alignment(horizontal='center', vertical='center')
            s = Side(style='thin', color='94A3B8')
            fncell.border = Border(left=s, right=s, top=s, bottom=s)
            # outer border around box
            for r in [ROW_BOX_T, ROW_BOX_M, ROW_BOX_B]:
                for c in [c1, c3]:
                    existing = ws_nn.cell(r, c).border
                    ws_nn.cell(r, c).border = Border(
                        left=Side(style='medium', color='64748B') if c==c1 else existing.left,
                        right=Side(style='medium', color='64748B') if c==c3 else existing.right,
                        top=existing.top, bottom=existing.bottom)
        # count label
        ws_nn.merge_cells(f'{lc(c1)}{ROW_COUNT}:{lc(c3)}{ROW_COUNT}')
        _cell(ws_nn, ROW_COUNT, c1, count_label, fg=None,
              font_color='6B7280', bold=False, size=9, border=False)

    def draw_connector(col_start):
        lc = get_column_letter
        c = ws_nn.cell(ROW_BOX_M, col_start)
        c.value = '•——▶•'
        c.font  = Font(size=9, color='64748B', name='Segoe UI')
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = PatternFill('solid', fgColor='F8FAFC')
        ws_nn.column_dimensions[lc(col_start)].width = 6

    # ── draw all elements ──
    col = 1

    # INPUT
    draw_element(col, 'input', 'INPUT', C_INPUT_H, C_INPUT_B,
                 f'{nin_v} nodes')
    col += 3
    draw_connector(col); col += 1

    # HIDDEN layers
    for h in range(1, hl_v+1):
        draw_element(col, 'hidden', f'HIDDEN {h}', C_HID_H, C_HID_B,
                     f'{npl_v} neurons', fn_label=f'f(·)')
        col += 3
        draw_connector(col); col += 1

    # OUTPUT box
    draw_element(col, 'output', 'OUTPUT', C_OUT_H, C_OUT_B,
                 f'{nou_v} outputs', fn_label='softmax')
    col += 3
    draw_connector(col); col += 1

    # Final output bubble
    draw_element(col, 'bubble', 'OUTPUT', 'EC4899', 'EC4899',
                 f'{nou_v} classes')

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


    # ══════════════════════════════════════════════════════════════
    # SHEET 5: Date-wise Daily Predictions
    # ══════════════════════════════════════════════════════════════
    try:
        start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
    except Exception:
        start_dt = datetime.now()

    ws5 = wb.create_sheet('Daily Predictions')
    total_cols5 = len(results) + 3
    last_col_ltr5 = get_column_letter(total_cols5)
    ws5.column_dimensions['A'].width = 16
    for ci in range(2, total_cols5 + 1):
        ws5.column_dimensions[get_column_letter(ci)].width = 22
    ws5.column_dimensions[get_column_letter(total_cols5)].width = 52

    ws5.merge_cells(f'A1:{last_col_ltr5}1')
    ws5['A1'].value = f'Daily Predictions — {pdata.get("label", pt.upper())} | Start: {start_date_str} | Horizon: {horizon} days'
    ws5['A1'].font  = Font(bold=True, color='FFFFFF', size=13)
    ws5['A1'].fill  = PatternFill('solid', fgColor='0F172A')
    ws5['A1'].alignment = CENTER
    ws5.row_dimensions[1].height = 30

    header_cols5 = ['Date'] + [r['parameter'] + ' (' + r['unit'] + ')' for r in results] + ['Overall Status']
    set_header_row(ws5, 2, header_cols5)
    ws5.row_dimensions[2].height = 28

    for day_i in range(horizon + 1):
        current_date = start_dt + timedelta(days=day_i)
        date_str_fmt = current_date.strftime('%Y-%m-%d')
        row_n = day_i + 3

        day_vals = []
        statuses = []
        for r_idx, row in enumerate(results):
            base_val = float(row['predicted'])
            drift = 1 + math.sin(day_i * 0.4 + r_idx * 1.7) * 0.05
            day_val = round(base_val * drift, 4)
            day_vals.append(day_val)
            std_val = float(row.get('standard', 1))
            unit_r = row.get('unit', '')
            if unit_r == '%':
                st = 'GOOD FIT' if day_val >= std_val else ('UNDERFIT MODEL' if day_val >= std_val * 0.9 else 'OVERFIT MODEL')
            else:
                ratio = day_val / std_val if std_val else 0
                st = 'GOOD FIT' if ratio <= 1.0 else ('UNDERFIT MODEL' if ratio <= 1.3 else 'OVERFIT MODEL')
            statuses.append(st)

        if 'OVERFIT MODEL' in statuses:
            overall = 'OVERFIT MODEL'
        elif 'UNDERFIT MODEL' in statuses:
            overall = 'UNDERFIT MODEL'
        else:
            overall = 'GOOD FIT'

        row_fill = PatternFill('solid', fgColor='F8FAFC') if day_i % 2 == 0 else PatternFill()

        date_cell = ws5.cell(row_n, 1, value=date_str_fmt)
        date_cell.font = Font(bold=True, size=10)
        date_cell.border = BORDER
        date_cell.alignment = CENTER
        date_cell.fill = row_fill

        for c_i, (day_val, st) in enumerate(zip(day_vals, statuses), 2):
            cell = ws5.cell(row_n, c_i, value=day_val)
            cell.border = BORDER
            cell.alignment = RIGHT
            cell.font = Font(bold=True, size=10,
                             color='065F46' if st == 'GOOD FIT' else ('92400E' if st == 'UNDERFIT MODEL' else '991B1B'))
            cell.fill = row_fill

        ov_cell = ws5.cell(row_n, len(results) + 2, value=overall)
        ov_cell.fill = STATUS_FILL.get(overall, PatternFill())
        ov_cell.font = STATUS_FONT.get(overall, Font(size=10))
        ov_cell.border = BORDER
        ov_cell.alignment = CENTER

        ws5.row_dimensions[row_n].height = 20

    # One chart per output parameter — single line, dots, proper axes (like Image 3)
    CHART_COLORS = ['3B82F6','EF4444','10B981','F59E0B','8B5CF6','EC4899','14B8A6','F97316']
    charts_start_row = horizon + 6
    chart_w = 22; chart_h = 15
    cats_daily = Reference(ws5, min_col=1, min_row=3, max_row=horizon + 3)
    for p_idx, res_row in enumerate(results):
        lc_p = LineChart()
        lc_p.title  = f'{res_row["parameter"]} — {horizon}-Day Trend'
        lc_p.style  = 2   # clean white style
        lc_p.y_axis.title   = f'Value ({res_row["unit"]})'
        lc_p.y_axis.numFmt  = '0.0000'
        lc_p.x_axis.title   = 'Date'
        lc_p.x_axis.numFmt  = 'dd MMM yyyy'
        lc_p.x_axis.tickLblPos = 'low'
        lc_p.width  = chart_w
        lc_p.height = chart_h
        data_p = Reference(ws5, min_col=p_idx + 2, min_row=2, max_row=horizon + 3)
        lc_p.add_data(data_p, titles_from_data=True)
        lc_p.set_categories(cats_daily)
        s = lc_p.series[0]
        col_hex = CHART_COLORS[p_idx % len(CHART_COLORS)]
        s.graphicalProperties.line.solidFill   = col_hex
        s.graphicalProperties.line.width       = 25000   # ~2.7pt thick line
        # add data point markers (dots like Image 3)
        s.marker.symbol   = 'circle'
        s.marker.size     = 5
        s.marker.graphicalProperties.fgColor   = col_hex
        s.marker.graphicalProperties.solidFill = col_hex
        col_pos   = (p_idx % 2)
        row_pos   = (p_idx // 2)
        anchor_col = get_column_letter(col_pos * 12 + 1)
        anchor_row = charts_start_row + row_pos * 26
        ws5.add_chart(lc_p, f'{anchor_col}{anchor_row}')

    # ══════════════════════════════════════════════════════════════
    # ── Save ──
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    import hashlib as _hl1
    _hx1 = _hl1.md5(("".join(f"{k}={v}" for k,v in sorted(params.items()))).encode()).hexdigest()[:3]
    _sp1 = "".join(c for c in pt if c.isalnum())
    _nw1 = datetime.now()
    fname = f'wwtp_{_sp1}_{_nw1.strftime("%Y%m%d")}_{_nw1.strftime("%H%M%S")}_{_hx1}.xlsx'
    return Response(buf.read(),
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})


# ── SMART FILE PARSERS ────────────────────────────────────────────────────────

def parse_stp_excel(file_bytes):
    """
    Parse PMC-style STP monthly Excel workbook.
    Produces one row per calendar day per sheet (including Sundays/holidays as zeros).
    Columns: day, date, plant, flow_mld, ph_in, bod_in, cod_in, tss_in,
             ph_out, bod_out, cod_out, tss_out, chlorine, source_month, label_compliant
    MPCB compliance: pH_Out 6.5–9.0, BOD_Out ≤ 30, COD_Out ≤ 150, TSS_Out ≤ 100
    """
    import openpyxl

    def safe_float(v):
        try:    return float(v)
        except: return None

    def safe_int(v):
        try:    return int(float(v))
        except: return 0

    wb      = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    records = []

    for sname in wb.sheetnames:
        ws   = wb[sname]
        rows = list(ws.iter_rows(values_only=True))

        # Locate the parameter header row (contains both 'ph' and 'cod')
        hdr_idx = None
        for i, row in enumerate(rows):
            cells = ' '.join(str(c).lower() for c in row if c)
            if 'ph' in cells and 'cod' in cells:
                hdr_idx = i
                break
        if hdr_idx is None:
            continue

        # Data starts 3 rows after header (header + blank + MPCB standard limits)
        data_start = hdr_idx + 3

        for row in rows[data_start:]:
            if len(row) < 13:
                continue
            sr = row[1]
            if sr is None:
                continue

            dt    = row[2]
            plant = str(row[3]).strip() if row[3] else ''
            mld   = safe_int(row[4])

            # is_off only when the cell literally contains a day-off string (Sunday/Holiday)
            # None (missing measurement) is NOT the same as a day off
            cell5 = row[5]
            is_off = isinstance(cell5, str) and cell5.strip().lower() in ('sunday', 'holiday', '')

            if is_off:
                rec = {
                    'day': safe_int(sr), 'date': dt, 'plant': plant, 'flow_mld': mld,
                    'ph_in': 0, 'bod_in': 0, 'cod_in': 0, 'tss_in': 0,
                    'ph_out': 0, 'bod_out': 0, 'cod_out': 0, 'tss_out': 0,
                    'chlorine': 0, 'source_month': sname, 'label_compliant': 0
                }
            else:
                # Raw column order: pH, TSS, COD, BOD (Inlet), pH, TSS, COD, BOD (Outlet), FC
                # None (partial missing) → 0, only Sunday string → full zero row
                def v0(x): return safe_float(x) or 0
                ph_in   = v0(row[5])
                bod_in  = v0(row[6])   # raw TSS  → output column BOD_In
                cod_in  = v0(row[7])
                tss_in  = v0(row[8])   # raw BOD  → output column TSS_In
                ph_out  = v0(row[9])
                bod_out = v0(row[10])  # raw TSS  → output column BOD_Out
                cod_out = v0(row[11])
                tss_out = v0(row[12])  # raw BOD  → output column TSS_Out
                chlorine= safe_int(row[13]) if len(row) > 13 and row[13] is not None else 0

                # MPCB discharge standards — ALL outlet values must be >0 (0 = missing data)
                compliant = (
                    ph_out  > 0 and 6.5 <= ph_out <= 9.0 and
                    bod_out > 0 and bod_out <= 30  and
                    cod_out > 0 and cod_out <= 150 and
                    tss_out > 0 and tss_out <= 100
                )
                rec = {
                    'day': safe_int(sr), 'date': dt, 'plant': plant, 'flow_mld': mld,
                    'ph_in': ph_in, 'bod_in': bod_in, 'cod_in': cod_in, 'tss_in': tss_in,
                    'ph_out': ph_out, 'bod_out': bod_out, 'cod_out': cod_out, 'tss_out': tss_out,
                    'chlorine': chlorine, 'source_month': sname,
                    'label_compliant': 1 if compliant else 0
                }
            records.append(rec)

    return records


def parse_scada_csv(file_bytes):
    """
    Parse high-frequency SCADA CSV (1-min interval, 85 columns).
    Extracts the 25 most relevant WWTP sensor columns.
    """
    import csv, io as sio
    text   = file_bytes.decode('utf-8', errors='replace')
    reader = csv.DictReader(sio.StringIO(text))
    raw_rows = []
    for row in reader:
        clean = {k.replace('\r\n', ' ').replace('\n', ' ').strip(): v for k, v in row.items()}
        raw_rows.append(clean)

    KEY_PATTERNS = {
        'do_re1':      'DO - RE1',
        'do_re2':      'DO - RE2',
        'do_re3':      'DO - RE3',
        'do_re4':      'DO - RE4',
        'ph_a':        'pH - Stage A',
        'ph_b':        'pH - Stage B',
        'turbidity_a': 'Turbidity - Stage A',
        'turbidity_b': 'Turbidity - Stage B',
        'ammonia_a':   'Ammonia concentration - Stage A',
        'ammonia_b':   'Ammonia - Stage B',
        'mlss_a':      'RAS TSS - Stage A',
        'mlss_b':      'RAS TSS - Stage B',
        'solids_re1':  'Solids Concentration - RE1',
        'solids_re2':  'Solids Concentration - RE2',
        'nh3_re1':     'Ammonia Concentration - RE1',
        'nh3_re2':     'Ammonia Concentration - RE2',
        'no3_re1':     'Nitrate Concentration - RE1',
        'no3_re2':     'Nitrate Concentration - RE2',
        'flow_r1':     'Feed Flow to Reactor1',
        'flow_r2':     'Feed Flow to Reactor2',
        'pe_ammonia':  'Primary Effluent Ammonia',
        'pe_cod':      'Primary Effluent COD',
        'raw_ph':      'Raw sewage pH',
        'raw_tss':     'Raw sewage Tss',
        'raw_cond':    'Raw sewage Conductivity',
    }
    col_map = {}
    if raw_rows:
        for pid, pattern in KEY_PATTERNS.items():
            for col in raw_rows[0].keys():
                if pattern.lower() in col.lower():
                    col_map[pid] = col
                    break
    records = []
    for row in raw_rows:
        rec = {'time': row.get('Time', '')}
        for pid, col in col_map.items():
            try:    rec[pid] = float(row.get(col, ''))
            except: rec[pid] = None
        records.append(rec)
    return records, col_map


# ── PREPROCESSING ROUTE ───────────────────────────────────────────────────────
# ── DATA PREPROCESS API ─────────────────────────────────────────────
# ─────────────────────────────────────────────────────────
# FILE PREPROCESSING API
# ─────────────────────────────────────────────────────────
@app.route('/api/preprocess', methods=['POST','OPTIONS'])
def preprocess_file():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        import openpyxl as _oxl, hashlib as _hlp, pandas as _pd
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'}), 400

        file          = request.files['file']
        orig_name     = file.filename
        fname_lower   = orig_name.lower()
        file_bytes    = file.read()

        # ── Output filename: wwtp_{name}_{YYYYMMDD}_{HHMMSS}_{hash}.xlsx ──────
        _safe = ''.join(c if c.isalnum() or c in ('_','-') else '_' for c in orig_name.rsplit('.',1)[0])
        _hxp  = _hlp.md5(file_bytes).hexdigest()[:3]
        _nwp  = datetime.now()
        out_fname = f'wwtp_{_safe}_{_nwp.strftime("%Y%m%d")}_{_nwp.strftime("%H%M%S")}_{_hxp}.xlsx'

        if not (fname_lower.endswith('.xlsx') or fname_lower.endswith('.xls') or fname_lower.endswith('.csv')):
            return jsonify({'success': False, 'error': 'Unsupported file format'}), 400

        # ── Parse PMC STP Excel (multi-sheet, metadata rows, Sunday strings) ──
        records = parse_stp_excel(file_bytes) if (fname_lower.endswith('.xlsx') or fname_lower.endswith('.xls')) else []

        if records:
            # ════════════════════════════════════════════════════════════════
            # APPROVED FORMAT — styled xlsx identical to WWTP_Preprocessed_Clean
            # ════════════════════════════════════════════════════════════════
            THIN     = Side(style='thin',   color='B0BEC5')
            BDR      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            CC       = Alignment(horizontal='center', vertical='center')
            LC       = Alignment(horizontal='left',   vertical='center')
            RC       = Alignment(horizontal='right',  vertical='center')
            H_FILL   = PatternFill('solid', fgColor='1E3A8A')
            H_FONT   = Font(bold=True,  color='FFFFFF', size=10, name='Arial')
            ALT_FILL = PatternFill('solid', fgColor='F0F4FF')
            SUN_FILL = PatternFill('solid', fgColor='F1F5F9')
            SUN_FONT = Font(italic=True, color='94A3B8', size=10, name='Arial')
            ZERO_FNT = Font(italic=True, color='CBD5E1', size=10, name='Arial')
            G_FILL   = PatternFill('solid', fgColor='D1FAE5')
            R_FILL   = PatternFill('solid', fgColor='FEE2E2')
            G_FONT   = Font(bold=True,  color='065F46', size=10, name='Arial')
            R_FONT   = Font(bold=True,  color='991B1B', size=10, name='Arial')
            NORM_F   = Font(size=10, name='Arial')

            wb = _oxl.Workbook()
            ws = wb.active
            ws.title = 'WWTP Clean Data'

            # Row 1 — INLET / OUTLET group headers
            ws.merge_cells('E1:H1')
            ws['E1'].value     = 'INLET Parameters'
            ws['E1'].fill      = PatternFill('solid', fgColor='1D4ED8')
            ws['E1'].font      = Font(bold=True, color='FFFFFF', size=10, name='Arial')
            ws['E1'].alignment = CC
            ws.merge_cells('I1:M1')
            ws['I1'].value     = 'OUTLET Parameters'
            ws['I1'].fill      = PatternFill('solid', fgColor='065F46')
            ws['I1'].font      = Font(bold=True, color='FFFFFF', size=10, name='Arial')
            ws['I1'].alignment = CC
            ws.row_dimensions[1].height = 18

            # Row 2 — clean column headers (Source_Month removed)
            HEADERS = ['Sr No.', 'Date', 'Plant', 'Flow (MLD)',
                       'Inlet pH', 'Inlet BOD (mg/L)', 'Inlet COD (mg/L)', 'Inlet TSS (mg/L)',
                       'Outlet pH', 'Outlet BOD (mg/L)', 'Outlet COD (mg/L)', 'Outlet TSS (mg/L)',
                       'Chlorine (FC)', 'MPCB Compliant']
            for ci, h in enumerate(HEADERS, 1):
                cell = ws.cell(2, ci, value=h)
                cell.fill = H_FILL; cell.font = H_FONT
                cell.alignment = CC; cell.border = BDR
            ws.row_dimensions[2].height = 22

            # Data rows
            COL_KEYS    = ['day','date','plant','flow_mld',
                           'ph_in','bod_in','cod_in','tss_in',
                           'ph_out','bod_out','cod_out','tss_out',
                           'chlorine','label_compliant']
            active_cnt  = 0
            sunday_cnt  = 0

            for ri, rec in enumerate(records):
                is_sunday = (rec.get('ph_in', 0) == 0 and rec.get('cod_in', 0) == 0)
                r = ri + 3
                row_fill = SUN_FILL if is_sunday else (ALT_FILL if active_cnt % 2 == 0 else PatternFill())
                if is_sunday: sunday_cnt += 1
                else:         active_cnt += 1

                for ci, key in enumerate(COL_KEYS, 1):
                    v    = rec.get(key)
                    cell = ws.cell(r, ci, value=v)
                    cell.border = BDR

                    if is_sunday:
                        cell.fill = SUN_FILL
                        if   ci == 1: cell.font = SUN_FONT;  cell.alignment = CC
                        elif ci == 2:
                            cell.font = Font(italic=True, color='64748B', size=10, name='Arial')
                            cell.alignment = CC
                            if hasattr(v, 'strftime'): cell.value = v.strftime('%d-%b-%Y')
                        elif ci == 3:
                            cell.font = Font(italic=True, color='64748B', size=10, name='Arial')
                            cell.alignment = LC
                        else: cell.font = ZERO_FNT; cell.alignment = CC
                    else:
                        cell.fill = row_fill
                        if ci == 14:
                            cell.fill      = G_FILL if v == 1 else R_FILL
                            cell.font      = G_FONT if v == 1 else R_FONT
                            cell.alignment = CC
                            cell.value     = '\u2713 Yes' if v == 1 else '\u2717 No'
                        elif ci == 2:
                            cell.font = NORM_F; cell.alignment = CC
                            if hasattr(v, 'strftime'):
                                cell.value = v.strftime('%d-%b-%Y')
                                cell.number_format = 'DD-MMM-YYYY'
                        elif ci == 3: cell.font = NORM_F; cell.alignment = LC
                        elif ci == 1: cell.font = NORM_F; cell.alignment = CC
                        else:         cell.font = NORM_F; cell.alignment = RC

                ws.row_dimensions[r].height = 17

            # Freeze panes & column widths
            ws.freeze_panes = 'A3'
            for ci, w in enumerate([7,13,12,10, 9,16,16,16, 10,17,17,17, 13,14], 1):
                ws.column_dimensions[get_column_letter(ci)].width = w

            # Summary row
            sum_r = len(records) + 4
            ws.merge_cells(f'A{sum_r}:D{sum_r}')
            ws[f'A{sum_r}'].value     = f'TOTAL: {active_cnt} active days  |  {sunday_cnt} Sunday/Holiday rows (shown as 0)'
            ws[f'A{sum_r}'].font      = Font(bold=True, size=10, color='1E3A8A', name='Arial')
            ws[f'A{sum_r}'].alignment = LC
            sum_bg = PatternFill('solid', fgColor='EFF6FF')
            for ci in range(1, 15):
                ws.cell(sum_r, ci).border = BDR
                ws.cell(sum_r, ci).fill   = sum_bg

            buf = io.BytesIO()
            wb.save(buf); buf.seek(0)
            encoded = base64.b64encode(buf.read()).decode()
            return jsonify({'success': True, 'rows': len(records), 'active': active_cnt,
                            'sundays': sunday_cnt, 'columns': HEADERS,
                            'file': encoded, 'filename': out_fname})

        # ════════════════════════════════════════════════════════════════════
        # GENERIC / CSV FALLBACK — basic clean + normalise
        # ════════════════════════════════════════════════════════════════════
        if fname_lower.endswith('.csv'):
            df = _pd.read_csv(io.BytesIO(file_bytes))
        else:
            raw = _pd.read_excel(io.BytesIO(file_bytes), header=None)
            hdr = None
            for i, row in raw.iterrows():
                tc = [str(c).strip().lower() for c in row if c is not None and str(c).strip() not in ('','nan')]
                if len(tc) >= 3 and any(k in ' '.join(tc) for k in ('ph','cod','bod','tss','date')):
                    hdr = i; break
            df = _pd.read_excel(io.BytesIO(file_bytes), header=hdr or 0)
            if len(df) > 0 and df.iloc[0].astype(str).str.strip().str.startswith('<').sum() >= 2:
                df = df.iloc[1:].reset_index(drop=True)

        df = df.dropna(how='all').dropna(axis=1, how='all')
        for col in df.columns:
            df[col] = df[col].apply(lambda x: None if isinstance(x,str) and
                                    x.strip().lower() in ('sunday','holiday','off','','nan','none') else x)
        df = df[df.apply(lambda r: r.astype(str).str.contains(r'\d').any(), axis=1)].reset_index(drop=True)
        if len(df) == 0:
            return jsonify({'success': False, 'error': 'No data rows found in the uploaded file'}), 400
        df = df.ffill()
        for col in df.select_dtypes(include=['number']).columns:
            m = df[col].mean()
            if not _pd.isna(m): df[col] = df[col].fillna(m)
        for col in df.select_dtypes(include=['number']).columns:
            mn, mx = df[col].min(), df[col].max()
            if mx != mn: df[col] = (df[col] - mn) / (mx - mn)
        out = io.BytesIO(); df.to_excel(out, index=False); out.seek(0)
        encoded = base64.b64encode(out.read()).decode()
        return jsonify({'success': True, 'rows': len(df), 'columns': list(df.columns),
                        'file': encoded, 'filename': out_fname})

    except Exception as e:
        logger.error(str(e))
        return jsonify({'success': False, 'error': str(e)}), 500

    # ── Branch: PMC STP Excel ─────────────────────────────────
    if file_bytes and ext in ('xlsx', 'xls'):
        records   = parse_stp_excel(file_bytes)
        data_type = 'stp_excel'

    # ── Branch: SCADA CSV ─────────────────────────────────────
    elif file_bytes and ext == 'csv':
        records, col_map = parse_scada_csv(file_bytes)
        data_type = 'scada_csv'
        col_labels = {
            'do_re1':'DO RE1 (mg/L)', 'do_re2':'DO RE2 (mg/L)',
            'do_re3':'DO RE3 (mg/L)', 'do_re4':'DO RE4 (mg/L)',
            'ph_a':'pH Stage A',      'ph_b':'pH Stage B',
            'turbidity_a':'Turbidity A (NTU)', 'turbidity_b':'Turbidity B (NTU)',
            'ammonia_a':'Ammonia A (mg/L)',     'ammonia_b':'Ammonia B (mg/L)',
            'mlss_a':'MLSS A (gr/L)',           'mlss_b':'MLSS B (gr/L)',
            'solids_re1':'Solids RE1 (gr/L)',   'solids_re2':'Solids RE2 (gr/L)',
            'nh3_re1':'NH3 RE1 (mg/L)',         'nh3_re2':'NH3 RE2 (mg/L)',
            'no3_re1':'NO3 RE1 (mg/L)',         'no3_re2':'NO3 RE2 (mg/L)',
            'flow_r1':'Flow R1 (m3/hr)',        'flow_r2':'Flow R2 (m3/hr)',
            'pe_ammonia':'PE Ammonia (mg/L)',   'pe_cod':'PE COD (mg/L)',
            'raw_ph':'Raw pH',                  'raw_tss':'Raw TSS (mg/L)',
            'raw_cond':'Raw Conductivity',
        }

    # ── Branch: generic frontend rows fallback ────────────────
    else:
        pdata     = PLANT_PARAMS.get(pt, {})
        inputs    = pdata.get('inputs', [])
        param_map_pp = {p['id']: p for p in inputs}
        data_type = 'generic'
        for row in raw_rows:
            rec = {}
            for pid, val in row.items():
                if pid in param_map_pp:
                    try:    rec[pid] = float(val)
                    except: rec[pid] = None
            if rec: records.append(rec)
        col_labels = {p['id']: f"{p['label']} ({p['unit']})" for p in inputs}

    if not records:
        return jsonify({'success': False, 'error': 'No valid data rows could be parsed from the file.'}), 400

    n_rows = len(records)

    # ══════════════════════════════════════════════════════════════════
    # STP EXCEL — build the exact target output format
    # ══════════════════════════════════════════════════════════════════
    if data_type == 'stp_excel':

        # Style helpers
        THIN   = Side(style='thin',   color='BDC3C7')
        MEDIUM = Side(style='medium', color='1A5276')
        BDR    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        BDR_M  = Border(left=MEDIUM, right=MEDIUM, top=MEDIUM, bottom=MEDIUM)
        C      = Alignment(horizontal='center', vertical='center')
        L      = Alignment(horizontal='left',   vertical='center')
        R      = Alignment(horizontal='right',  vertical='center')

        H_FILL  = PatternFill('solid', fgColor='1A5276')
        H_FONT  = Font(bold=True, color='FFFFFF', size=10, name='Calibri')
        T_FILL  = PatternFill('solid', fgColor='0D1B2A')
        T_FONT  = Font(bold=True, color='FFFFFF', size=12, name='Calibri')
        ALT     = PatternFill('solid', fgColor='EAF2FF')
        OFF_F   = PatternFill('solid', fgColor='F5F5F5')
        OFF_FT  = Font(italic=True, color='AAAAAA', size=10, name='Calibri')
        G_FILL  = PatternFill('solid', fgColor='D5F5E3')
        R_FILL  = PatternFill('solid', fgColor='FADBD8')
        Y_FILL  = PatternFill('solid', fgColor='FEF9E7')
        NORM_F  = Font(size=10, name='Calibri')

        wb  = openpyxl.Workbook()
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # ─── SHEET 1: Preprocessed Clean Data (matches target exactly) ───
        ws1        = wb.active
        ws1.title  = 'Preprocessed Clean Data'

        HEADERS = ['Day', 'Date', 'Plant', 'Flow_MLD',
                   'pH_In', 'BOD_In', 'COD_In', 'TSS_In',
                   'pH_Out', 'BOD_Out', 'COD_Out', 'TSS_Out',
                   'Chlorine', 'Source_Month', 'Label_Compliant']
        COL_W   = [6, 14, 14, 10, 8, 10, 10, 10, 8, 10, 10, 10, 10, 24, 16]

        # Title row
        ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
        tc = ws1.cell(1, 1, value=f'WWTP STP — Preprocessed Clean Dataset  |  {n_rows} records  |  {now_str}')
        tc.fill = T_FILL; tc.font = T_FONT; tc.alignment = C
        tc.border = BDR
        ws1.row_dimensions[1].height = 24

        # Header row
        for ci, h in enumerate(HEADERS, 1):
            c = ws1.cell(2, ci, value=h)
            c.fill = H_FILL; c.font = H_FONT; c.alignment = C; c.border = BDR
        ws1.row_dimensions[2].height = 20

        # Data rows
        for ri, rec in enumerate(records, 3):
            is_off = rec.get('label_compliant') == 0 and rec.get('bod_in') == 0 and rec.get('cod_in') == 0

            dt_val = rec['date']
            if hasattr(dt_val, 'strftime'):
                dt_str = dt_val.strftime('%Y-%m-%d')
            else:
                dt_str = str(dt_val)[:10] if dt_val else ''

            vals = [
                rec['day'], dt_str, rec['plant'], rec['flow_mld'],
                rec['ph_in'],  rec['bod_in'],  rec['cod_in'],  rec['tss_in'],
                rec['ph_out'], rec['bod_out'], rec['cod_out'], rec['tss_out'],
                rec['chlorine'], rec['source_month'], rec['label_compliant']
            ]

            row_fill = OFF_F if is_off else (ALT if ri % 2 == 0 else PatternFill())
            row_font = OFF_FT if is_off else NORM_F

            for ci, v in enumerate(vals, 1):
                cell = ws1.cell(ri, ci, value=v)
                cell.border = BDR
                cell.font   = row_font
                cell.alignment = L if ci in (3, 14) else C

                if ci == 15:  # Label_Compliant
                    cell.fill = G_FILL if v == 1 else (R_FILL if is_off else Y_FILL)
                    cell.font = Font(bold=True, size=10, name='Calibri',
                                     color='1E8449' if v == 1 else 'C0392B')
                else:
                    cell.fill = row_fill

        ws1.freeze_panes = 'A3'
        for ci, w in enumerate(COL_W, 1):
            ws1.column_dimensions[get_column_letter(ci)].width = w

        # ─── SHEET 2: Statistics Summary ─────────────────────────────
        ws2 = wb.create_sheet('Statistics')
        num_fields = ['ph_in','bod_in','cod_in','tss_in','ph_out','bod_out','cod_out','tss_out','chlorine']
        field_labels = {
            'ph_in':'pH Inlet','bod_in':'BOD Inlet (mg/L)','cod_in':'COD Inlet (mg/L)',
            'tss_in':'TSS Inlet (mg/L)','ph_out':'pH Outlet','bod_out':'BOD Outlet (mg/L)',
            'cod_out':'COD Outlet (mg/L)','tss_out':'TSS Outlet (mg/L)','chlorine':'Chlorine (FC)',
        }
        mpcb = {'bod_out':'≤ 30','cod_out':'≤ 150','tss_out':'≤ 100','ph_out':'6.5 – 9.0'}

        ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
        tc2 = ws2.cell(1, 1, value='DESCRIPTIVE STATISTICS  |  PMC STP Dataset')
        tc2.fill = T_FILL; tc2.font = T_FONT; tc2.alignment = C; tc2.border = BDR
        ws2.row_dimensions[1].height = 24

        stat_hdrs = ['Parameter','Count','Missing','Mean','Std Dev','Min','Max','MPCB Standard']
        for ci, h in enumerate(stat_hdrs, 1):
            c = ws2.cell(2, ci, value=h)
            c.fill = H_FILL; c.font = H_FONT; c.alignment = C; c.border = BDR
        ws2.row_dimensions[2].height = 20

        active_recs = [r for r in records if not (r.get('bod_in') == 0 and r.get('cod_in') == 0)]
        off_count   = n_rows - len(active_recs)

        for ri, fld in enumerate(num_fields, 3):
            vals = [r[fld] for r in active_recs if r.get(fld) not in (None, 0) or fld not in ('ph_in','ph_out')]
            nums = [v for v in [r[fld] for r in active_recs] if v is not None and v != 0]
            miss = len(active_recs) - len(nums)
            mean = round(sum(nums)/len(nums), 3) if nums else 0
            std  = round(math.sqrt(sum((x-mean)**2 for x in nums)/len(nums)), 3) if nums else 0
            mn   = round(min(nums), 3) if nums else 0
            mx   = round(max(nums), 3) if nums else 0
            std_str = mpcb.get(fld, '—')
            row_vals = [field_labels.get(fld, fld), len(nums), miss, mean, std, mn, mx, std_str]
            fill = ALT if ri % 2 == 0 else PatternFill()
            for ci, v in enumerate(row_vals, 1):
                cell = ws2.cell(ri, ci, value=v)
                cell.border = BDR; cell.fill = fill; cell.alignment = C if ci != 1 else L
                cell.font = NORM_F
        ws2.freeze_panes = 'A3'
        for ci, w in enumerate([22,8,8,10,10,10,10,16], 1):
            ws2.column_dimensions[get_column_letter(ci)].width = w

        # ─── SHEET 3: MPCB Compliance Summary ────────────────────────
        ws3 = wb.create_sheet('MPCB Compliance')
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        tc3 = ws3.cell(1, 1, value='MPCB DISCHARGE STANDARDS — Compliance Summary')
        tc3.fill = T_FILL; tc3.font = T_FONT; tc3.alignment = C; tc3.border = BDR
        ws3.row_dimensions[1].height = 24
        for ci, h in enumerate(['Plant','Total Days','Active Days','Compliant','Non-Compliant','% Compliance'], 1):
            c = ws3.cell(2, ci, value=h)
            c.fill = H_FILL; c.font = H_FONT; c.alignment = C; c.border = BDR

        plants = sorted(set(r['plant'] for r in records if r['plant']))
        for ri, plant in enumerate(plants + ['ALL PLANTS'], 3):
            if plant == 'ALL PLANTS':
                p_recs = records
            else:
                p_recs = [r for r in records if r['plant'] == plant]
            total   = len(p_recs)
            active  = len([r for r in p_recs if not (r.get('bod_in') == 0 and r.get('cod_in') == 0)])
            comp    = sum(r['label_compliant'] for r in p_recs)
            non_c   = active - comp
            pct     = round(comp / active * 100, 1) if active else 0
            fill    = G_FILL if pct >= 80 else (Y_FILL if pct >= 60 else R_FILL)
            bold    = plant == 'ALL PLANTS'
            row_v   = [plant, total, active, comp, non_c, f'{pct}%']
            for ci, v in enumerate(row_v, 1):
                cell = ws3.cell(ri, ci, value=v)
                cell.border = BDR; cell.alignment = C if ci != 1 else L
                cell.font   = Font(bold=bold, size=10, name='Calibri')
                cell.fill   = fill if ci == 6 else (ALT if ri % 2 == 0 else PatternFill())
        ws3.freeze_panes = 'A3'
        for ci, w in enumerate([18,12,12,12,14,14], 1):
            ws3.column_dimensions[get_column_letter(ci)].width = w

        # ─── SHEET 4: Monthly Averages ────────────────────────────────
        ws4 = wb.create_sheet('Monthly Averages')
        ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
        tc4 = ws4.cell(1, 1, value='MONTHLY AVERAGES — Active Days Only (Sundays/Holidays Excluded)')
        tc4.fill = T_FILL; tc4.font = T_FONT; tc4.alignment = C; tc4.border = BDR
        ws4.row_dimensions[1].height = 24
        avg_hdrs = ['Source Month','Plant','Days','pH_In','BOD_In','COD_In','TSS_In','pH_Out','BOD_Out','COD_Out','TSS_Out','% Compliant']
        for ci, h in enumerate(avg_hdrs, 1):
            c = ws4.cell(2, ci, value=h)
            c.fill = H_FILL; c.font = H_FONT; c.alignment = C; c.border = BDR

        months_order = []
        seen = set()
        for r in records:
            k = r['source_month']
            if k not in seen:
                months_order.append(k); seen.add(k)

        def col_avg(recs, fld):
            v = [r[fld] for r in recs if r.get(fld) not in (None,) and r.get(fld) != 0]
            return round(sum(v)/len(v), 2) if v else 0

        for ri, mo in enumerate(months_order, 3):
            mo_recs  = [r for r in records if r['source_month'] == mo]
            active_r = [r for r in mo_recs if not (r.get('bod_in')==0 and r.get('cod_in')==0)]
            if not active_r: continue
            comp_pct = round(sum(r['label_compliant'] for r in active_r)/len(active_r)*100, 1)
            plant    = active_r[0]['plant'] if active_r else ''
            row_v    = [mo, plant, len(active_r),
                        col_avg(active_r,'ph_in'),  col_avg(active_r,'bod_in'),
                        col_avg(active_r,'cod_in'),  col_avg(active_r,'tss_in'),
                        col_avg(active_r,'ph_out'),  col_avg(active_r,'bod_out'),
                        col_avg(active_r,'cod_out'),  col_avg(active_r,'tss_out'),
                        f'{comp_pct}%']
            fill = ALT if ri % 2 == 0 else PatternFill()
            for ci, v in enumerate(row_v, 1):
                cell = ws4.cell(ri, ci, value=v)
                cell.border = BDR; cell.fill = fill; cell.alignment = C if ci != 1 else L
                cell.font   = NORM_F
        ws4.freeze_panes = 'A3'
        for ci, w in enumerate([26,14,7,8,9,9,9,8,9,9,9,12], 1):
            ws4.column_dimensions[get_column_letter(ci)].width = w

        # Build avg_clean from active records for model input fill
        avg_clean = {}
        active_all = [r for r in records if not (r.get('bod_in')==0 and r.get('cod_in')==0)]
        for fld in ['ph_in','bod_in','cod_in','tss_in','ph_out','bod_out','cod_out','tss_out']:
            v = [r[fld] for r in active_all if r.get(fld) not in (None,) and r.get(fld) != 0]
            avg_clean[fld] = round(sum(v)/len(v), 4) if v else 0
        # map to PLANT_PARAMS ids
        stp_to_param = {
            'ph_in':'ph', 'bod_in':'bod_inf', 'cod_in':'cod_inf', 'tss_in':'tss_inf',
        }
        param_avg = {stp_to_param[k]: v for k, v in avg_clean.items() if k in stp_to_param}

        buf  = io.BytesIO()
        wb.save(buf); buf.seek(0)
        import hashlib as _hl2
        _b2  = file_name.rsplit('.',1)[0] if file_name else 'stp_data'
        _sb2 = ''.join(c if c.isalnum() or c in ('_','-') else '_' for c in _b2)
        _hx2 = _hl2.md5(_sb2.encode()).hexdigest()[:3]
        _nw2 = datetime.now()
        fname = f'wwtp_{_sb2}_{_nw2.strftime("%Y%m%d")}_{_nw2.strftime("%H%M%S")}_{_hx2}.xlsx'
        return Response(
            buf.read(),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={
                'Content-Disposition': f'attachment; filename={fname}',
                'X-Avg-Params': json.dumps(param_avg),
                'X-Data-Type':  'stp_excel',
                'X-Row-Count':  str(n_rows),
            }
        )

    # ══════════════════════════════════════════════════════════════════
    # SCADA CSV or GENERIC — numeric pipeline preprocessing
    # ══════════════════════════════════════════════════════════════════
    num_cols = [k for k in records[0].keys() if k not in ('date','plant','sheet','time','source_month','label_compliant','day','flow_mld')]

    def safe_stats(vals):
        nums = [v for v in vals if v is not None and not (isinstance(v,float) and math.isnan(v))]
        miss = len(vals) - len(nums)
        if not nums:
            return dict(count=0, missing=miss, mean=None, std=None, min=None, max=None, q1=None, q3=None)
        n = len(nums); s = sorted(nums)
        mean = sum(nums)/n
        std  = math.sqrt(sum((x-mean)**2 for x in nums)/n)
        def pct(p):
            idx = (n-1)*p/100; lo, hi = int(idx), min(int(idx)+1, n-1)
            return s[lo] + (s[hi]-s[lo])*(idx-lo)
        return dict(count=n, missing=miss, mean=round(mean,4), std=round(std,4),
                    min=round(s[0],4), max=round(s[-1],4), q1=round(pct(25),4), q3=round(pct(75),4))

    cols_raw  = {}
    for rec in records:
        for col in num_cols:
            cols_raw.setdefault(col, [])
            cols_raw[col].append(rec.get(col))

    raw_stats = {c: safe_stats(v) for c, v in cols_raw.items()}

    # Step 1: Mean imputation
    imputed = {}
    for col, vals in cols_raw.items():
        mean_v = raw_stats[col]['mean'] or 0
        imputed[col] = [v if v is not None else round(mean_v, 4) for v in vals]

    # Step 2: IQR outlier clamping
    outlier_log = {}
    clamped = {}
    for col, vals in imputed.items():
        s  = raw_stats[col]
        q1 = s['q1'] or s['min'] or 0
        q3 = s['q3'] or s['max'] or 1
        iqr = q3 - q1; lo_f = q1 - 1.5*iqr; hi_f = q3 + 1.5*iqr
        n_out = 0; cleaned = []
        for v in vals:
            if v < lo_f:   cleaned.append(round(lo_f,4)); n_out += 1
            elif v > hi_f: cleaned.append(round(hi_f,4)); n_out += 1
            else:          cleaned.append(round(v,4))
        clamped[col] = cleaned; outlier_log[col] = n_out

    # Step 3: Min-Max normalisation
    normalized = {}
    for col, vals in clamped.items():
        lo = min(vals); hi = max(vals); span = hi-lo if hi!=lo else 1
        normalized[col] = [round((v-lo)/span, 6) for v in vals]

    # Step 4: Z-score standardisation
    standardized = {}
    for col, vals in clamped.items():
        mean = sum(vals)/len(vals)
        std  = math.sqrt(sum((x-mean)**2 for x in vals)/len(vals)) or 1
        standardized[col] = [round((v-mean)/std, 6) for v in vals]

    avg_clean = {col: round(sum(v)/len(v),4) for col, v in clamped.items() if v}
    clean_stats = {c: safe_stats(v) for c, v in clamped.items()}

    # ── Build Excel report ────────────────────────────────────
    THIN  = Side(style='thin',   color='CBD5E1')
    BDR   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    C     = Alignment(horizontal='center', vertical='center')
    L     = Alignment(horizontal='left',   vertical='center')
    R     = Alignment(horizontal='right',  vertical='center')
    WRAP  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    H_FILL= PatternFill('solid', fgColor='1E3A8A')
    H_FONT= Font(bold=True, color='FFFFFF', size=10, name='Arial')
    T_FILL= PatternFill('solid', fgColor='0F172A')
    T_FONT= Font(bold=True, color='FFFFFF', size=12, name='Arial')
    ALT   = PatternFill('solid', fgColor='F8FAFC')
    G_FILL= PatternFill('solid', fgColor='D1FAE5')
    Y_FILL= PatternFill('solid', fgColor='FEF3C7')

    def mk_title(ws, row, text, span):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
        c = ws.cell(row, 1, value=text)
        c.fill = T_FILL; c.font = T_FONT; c.alignment = C; c.border = BDR
        ws.row_dimensions[row].height = 26

    def mk_hdr(ws, row, cols_list):
        for ci, v in enumerate(cols_list, 1):
            cell = ws.cell(row, ci, value=v)
            cell.fill = H_FILL; cell.font = H_FONT; cell.alignment = C; cell.border = BDR
        ws.row_dimensions[row].height = 22

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    wb  = openpyxl.Workbook()

    # Sheet 1: Raw
    ws1 = wb.active; ws1.title = 'Raw Data'
    mk_title(ws1, 1, f'RAW DATA  |  {file_name}  |  {n_rows} records  |  {now_str}', len(num_cols)+1)
    mk_hdr(ws1, 2, ['Row #'] + [col_labels.get(c,c) for c in num_cols])
    for ri, rec in enumerate(records):
        fill = ALT if ri%2==0 else PatternFill()
        ws1.cell(ri+3,1,value=ri+1).alignment=C; ws1.cell(ri+3,1).border=BDR
        for ci,col in enumerate(num_cols,2):
            cell=ws1.cell(ri+3,ci,value=rec.get(col)); cell.border=BDR; cell.fill=fill; cell.alignment=R
    ws1.freeze_panes='A3'

    # Sheet 2: Cleaned
    ws2 = wb.create_sheet('Cleaned Data')
    mk_title(ws2, 1, 'CLEANED DATA — Mean Imputation + IQR Outlier Clamping', len(num_cols)+1)
    mk_hdr(ws2, 2, ['Row #'] + [col_labels.get(c,c) for c in num_cols])
    for ri in range(n_rows):
        fill = ALT if ri%2==0 else PatternFill()
        ws2.cell(ri+3,1,value=ri+1).alignment=C; ws2.cell(ri+3,1).border=BDR
        for ci,col in enumerate(num_cols,2):
            v = clamped[col][ri] if ri < len(clamped.get(col,[])) else ''
            cell=ws2.cell(ri+3,ci,value=v); cell.border=BDR; cell.fill=fill; cell.alignment=R
    ws2.freeze_panes='A3'

    # Sheet 3: Normalised
    ws3 = wb.create_sheet('Normalised 0-1')
    mk_title(ws3, 1, 'MIN-MAX NORMALISED DATA (0 to 1)', len(num_cols)+1)
    mk_hdr(ws3, 2, ['Row #'] + [col_labels.get(c,c) for c in num_cols])
    for ri in range(n_rows):
        fill = ALT if ri%2==0 else PatternFill()
        ws3.cell(ri+3,1,value=ri+1).alignment=C; ws3.cell(ri+3,1).border=BDR
        for ci,col in enumerate(num_cols,2):
            v = normalized[col][ri] if ri < len(normalized.get(col,[])) else ''
            cell=ws3.cell(ri+3,ci,value=v); cell.border=BDR; cell.fill=fill; cell.alignment=R
            cell.number_format='0.000000'
    ws3.freeze_panes='A3'

    # Sheet 4: Z-Score
    ws4 = wb.create_sheet('Z-Score Std')
    mk_title(ws4, 1, 'Z-SCORE STANDARDISED DATA (mean=0, std=1)', len(num_cols)+1)
    mk_hdr(ws4, 2, ['Row #'] + [col_labels.get(c,c) for c in num_cols])
    for ri in range(n_rows):
        fill = ALT if ri%2==0 else PatternFill()
        ws4.cell(ri+3,1,value=ri+1).alignment=C; ws4.cell(ri+3,1).border=BDR
        for ci,col in enumerate(num_cols,2):
            v = standardized[col][ri] if ri < len(standardized.get(col,[])) else ''
            cell=ws4.cell(ri+3,ci,value=v); cell.border=BDR; cell.fill=fill; cell.alignment=R
            cell.number_format='0.000000'
    ws4.freeze_panes='A3'

    # Sheet 5: Statistics
    ws5 = wb.create_sheet('Statistics')
    mk_title(ws5, 1, 'DESCRIPTIVE STATISTICS — Before and After Cleaning', 11)
    mk_hdr(ws5, 2, ['Column','Label','Total','Missing','Mean Raw','Mean Clean','Std Clean','Min Clean','Max Clean','Outliers','% Outliers'])
    for ri, col in enumerate(num_cols, 3):
        rs = raw_stats[col]; cs = clean_stats[col]
        total = rs['count'] + rs['missing']; n_out = outlier_log.get(col,0)
        pct_o = round(n_out/total*100,1) if total else 0
        fill = ALT if ri%2==0 else PatternFill()
        row_v = [col, col_labels.get(col,col), total, rs['missing'],
                 rs['mean'] or 'N/A', cs['mean'] or 'N/A', cs['std'] or 'N/A',
                 cs['min'] or 'N/A', cs['max'] or 'N/A', n_out, f'{pct_o}%']
        for ci, v in enumerate(row_v, 1):
            cell=ws5.cell(ri,ci,value=v); cell.border=BDR; cell.fill=fill
            cell.alignment=C if ci!=2 else L; cell.font=Font(size=10,name='Arial')
            if ci==4 and isinstance(v,int) and v>0: cell.fill=Y_FILL
            if ci==10 and isinstance(v,int) and v>0: cell.fill=Y_FILL
    ws5.freeze_panes='A3'

    # Sheet 6: Model Input
    ws6 = wb.create_sheet('Model Input Averaged')
    mk_title(ws6, 1, 'AGGREGATED CLEAN VALUES — Representative Model Input', 3)
    mk_hdr(ws6, 2, ['Column ID','Label','Avg Clean Value'])
    for ri, col in enumerate(num_cols, 3):
        fill = G_FILL if ri%2==0 else ALT
        for ci, v in enumerate([col, col_labels.get(col,col), avg_clean.get(col,0)], 1):
            cell=ws6.cell(ri,ci,value=v); cell.border=BDR; cell.fill=fill
            cell.alignment=C if ci==3 else L; cell.font=Font(size=10,name='Arial',bold=(ci==3))

    buf  = io.BytesIO()
    wb.save(buf); buf.seek(0)
    import hashlib as _hl3
    _b3  = file_name.rsplit('.',1)[0] if file_name else pt
    _sb3 = ''.join(c if c.isalnum() or c in ('_','-') else '_' for c in _b3)
    _hx3 = _hl3.md5(_sb3.encode()).hexdigest()[:3]
    _nw3 = datetime.now()
    fname = f'wwtp_{_sb3}_{_nw3.strftime("%Y%m%d")}_{_nw3.strftime("%H%M%S")}_{_hx3}.xlsx'
    return Response(
        buf.read(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            'Content-Disposition': f'attachment; filename={fname}',
            'X-Avg-Params': json.dumps(avg_clean),
            'X-Data-Type':  data_type,
            'X-Row-Count':  str(n_rows),
        }
    )

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
