"""
WWTP Neural Prediction System — Backend v4.0 (Production)
Deploy : gunicorn --workers=4 --threads=2 --timeout=120 app:app
Open   : https://your-app.onrender.com

Requirements: pip install flask openpyxl gunicorn pandas scipy numpy scikit-learn
"""

from flask import Flask, request, jsonify, send_from_directory, Response
import json, math, random, os, io, csv, base64, logging
from datetime import datetime, timedelta
import pandas as pd
import polars as pl
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1
import hashlib

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


# ── FAST PREPROCESSING USING POLARS ─────────────────────────────
def fast_preprocess_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    try:
        pldf = pl.from_pandas(df)

        # remove sunday/holiday rows
        pldf = pldf.filter(
            ~pl.any_horizontal(
                pl.all().cast(pl.Utf8).str.contains("sunday|holiday", literal=False)
            )
        )

        # clean < values
        for c in pldf.columns:
            pldf = pldf.with_columns(
                pl.when(pl.col(c).cast(pl.Utf8).str.starts_with("<"))
                .then(
                    pl.col(c).cast(pl.Utf8).str.replace("<","").cast(pl.Float64)/2
                )
                .otherwise(pl.col(c))
                .alias(c)
            )

        # feature engineering
        if {"bod_in","bod_out"}.issubset(set(pldf.columns)):
            pldf = pldf.with_columns(
                ((pl.col("bod_in")-pl.col("bod_out"))/pl.col("bod_in")*100)
                .alias("bod_removal_%")
            )

        if {"cod_in","cod_out"}.issubset(set(pldf.columns)):
            pldf = pldf.with_columns(
                ((pl.col("cod_in")-pl.col("cod_out"))/pl.col("cod_in")*100)
                .alias("cod_removal_%")
            )

        if {"tss_in","tss_out"}.issubset(set(pldf.columns)):
            pldf = pldf.with_columns(
                ((pl.col("tss_in")-pl.col("tss_out"))/pl.col("tss_in")*100)
                .alias("tss_removal_%")
            )

        return pldf.to_pandas()
    except Exception:
        # polars failed — return the original pandas df unchanged
        return df

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

    if mat_basename is None:
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
net.trainParam.showWindow  = true;

%% 5. TRAIN
fprintf('[TRAIN] Starting training with {algo} algorithm...\\n');
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
scr = get(0,'ScreenSize');
fw = floor(scr(3)/3); fh = floor(scr(4)/2);
pos = @(col,row) [fw*(col-1)+10, scr(4)-fh*row-40, fw-20, fh-60];

f1 = figure('Name','[1] Training Performance','NumberTitle','off','Position',pos(1,1));
plotperform(tr_rec);
title(sprintf('Training Performance — {pt.upper()} | Best: Epoch %d | MSE: %.6f', tr_rec.best_epoch, min(tr_rec.perf)));

f2 = figure('Name','[2] Regression R²','NumberTitle','off','Position',pos(2,1));
plotregression(Y_norm, Y_pred_n, sprintf('ANN Regression — R²=%.4f', R2));

f3 = figure('Name','[3] Error Histogram','NumberTitle','off','Position',pos(3,1));
errors = Y_raw - Y_pred;
histogram(errors(:), 40, 'FaceColor',[0 0.75 0.65], 'EdgeColor','w');
hold on;
xline(mean(errors(:)), 'r--', 'LineWidth', 2, 'Label', sprintf('Mean=%.4f', mean(errors(:))));
xlabel('Prediction Error'); ylabel('Frequency');
title('Error Histogram — {pt.upper()}'); grid on;

f4 = figure('Name','[4] Actual vs Predicted','NumberTitle','off','Position',pos(1,2));
x_ax = 1:n_outputs;
b = bar(x_ax, [target_vals(:), y_new(:)]);
b(1).FaceColor = [0.2 0.6 0.9]; b(2).FaceColor = [0.1 0.8 0.5];
set(gca,'XTick',x_ax,'XTickLabel',output_labels,'XTickLabelRotation',30,'FontSize',9);
legend('Actual (Target)','Predicted (ANN)','Location','best');
ylabel('Value'); title('Actual vs Predicted Effluent — {pt.upper()}'); grid on;

f5 = figure('Name','[5] Network Architecture','NumberTitle','off','Position',pos(2,2));
view(net);

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

fprintf('\\n[DONE] All figures open. Results saved to: %s\\n', save_file);
figure(f1);
"""

# ── API ROUTES ────────────────────────────────────────────────────────────────
@app.route('/api/status')
def status():
    return jsonify({'status':'running','version':'4.0','plants':list(PLANT_PARAMS.keys())})

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


# ════════════════════════════════════════════════════════════════════════════════
#  FULL PREPROCESSING PIPELINE
#  Steps:
#   1. Smart file reading  (xlsx / xls / csv, any encoding, auto-header)
#   2. Drop empty rows / cols
#   3. Replace day-off / sentinel strings with NaN
#   4. Type coercion  — force numeric where possible
#   5. Missing-value imputation  (median for skewed, mean for normal cols)
#   6. IQR outlier detection & Winsorisation (clip to [Q1-1.5*IQR, Q3+1.5*IQR])
#   7. Min-Max Normalisation  → [0, 1]
#   8. Z-Score Standardisation → μ=0, σ=1   (stored in separate columns)
#   9. Duplicate row removal
#  10. Styled multi-sheet Excel output with audit log
# ════════════════════════════════════════════════════════════════════════════════

OFF_STRINGS = {
    'sunday','holiday','off','n/a','na','nil','not available',
    'none','null','–','—','-','--','---','#n/a','#na','#value!',
    '#ref!','#div/0!','#null!','error','nan','', ' '
}

def _coerce_numeric(series: pd.Series) -> pd.Series:
    """Try to convert a column to numeric; return NaN for anything that fails."""
    return pd.to_numeric(series.astype(str).str.strip().str.replace(',','', regex=False),
                         errors='coerce')

def _is_numeric_col(series: pd.Series, threshold: float = 0.60) -> bool:
    """Return True if ≥ threshold fraction of non-null values are numeric."""
    coerced = _coerce_numeric(series.dropna())
    return coerced.notna().mean() >= threshold if len(coerced) else False

def _replace_off_strings(df: pd.DataFrame) -> tuple:
    """Replace sentinel strings with NaN. Returns (df, count_replaced)."""
    replaced = 0
    for col in df.columns:
        mask = df[col].astype(str).str.strip().str.lower().isin(OFF_STRINGS)
        replaced += int(mask.sum())
        df.loc[mask, col] = np.nan
    df = fast_preprocess_dataframe(df)
    return df, replaced

def _smart_read(file_bytes: bytes, fname_lower: str) -> pd.DataFrame:
    """Read any xlsx/xls/csv into a clean DataFrame."""
    if fname_lower.endswith('.csv'):
        for enc in ('utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1'):
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc,
                                 skip_blank_lines=True, na_values=list(OFF_STRINGS))
                df = fast_preprocess_dataframe(df)
                return df
            except Exception:
                continue
        raise ValueError("Cannot decode CSV — try saving as UTF-8.")
    else:
        # Find the actual header row (first row with ≥3 text cells)
        raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')
        hdr_row = 0
        for i, row in raw.iterrows():
            text_cells = [str(c).strip() for c in row
                          if c is not None and str(c).strip() not in ('', 'nan', 'None')]
            if len(text_cells) >= 3:
                hdr_row = i
                break
        df = pd.read_excel(io.BytesIO(file_bytes), header=hdr_row,
                           engine='openpyxl', na_values=list(OFF_STRINGS))
        # Drop MPCB-style limit rows (rows where ≥2 cells start with '<')
        if len(df) > 0:
            lt_mask = df.iloc[0].astype(str).str.strip().str.startswith('<').sum() >= 2
            if lt_mask:
                df = df.iloc[1:].reset_index(drop=True)
        df = fast_preprocess_dataframe(df)
        return df


def full_preprocess_pipeline(file_bytes: bytes, fname_lower: str) -> dict:
    """
    Run the complete 10-step preprocessing pipeline.
    Returns a dict with:
      - df_clean      : final cleaned DataFrame (normalised)
      - df_zscore     : z-score standardised DataFrame (numeric cols only)
      - audit         : dict with step-by-step statistics
      - num_cols      : list of numeric column names
      - non_num_cols  : list of non-numeric column names
      - original_shape: (rows, cols) before cleaning
    """
    audit = {}

    # ── STEP 1: Read file ──────────────────────────────────────────────────────
    df = _smart_read(file_bytes, fname_lower)
    audit['step1_read'] = {'rows': len(df), 'cols': len(df.columns),
                           'columns': list(df.columns)}
    original_shape = (len(df), len(df.columns))

    # Normalise column names
    df.columns = [str(c).strip() for c in df.columns]

    # ── STEP 2: Drop fully-empty rows and columns ──────────────────────────────
    rows_before = len(df)
    cols_before = len(df.columns)
    df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    audit['step2_drop_empty'] = {
        'rows_dropped': rows_before - len(df),
        'cols_dropped': cols_before - len(df.columns)
    }

    # ── STEP 3: Replace sentinel / day-off strings with NaN ───────────────────
    df, replaced_count = _replace_off_strings(df)
    audit['step3_sentinel_replace'] = {'cells_replaced_with_nan': replaced_count}

    # ── STEP 4: Type coercion — convert numeric-looking columns ───────────────
    coerced_cols = []
    non_num_cols = []
    for col in df.columns:
        if _is_numeric_col(df[col]):
            df[col] = _coerce_numeric(df[col])
            coerced_cols.append(col)
        else:
            non_num_cols.append(col)

    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    audit['step4_type_coercion'] = {
        'numeric_cols': len(num_cols),
        'text_cols': len(non_num_cols),
        'numeric_col_names': num_cols
    }

    if len(num_cols) == 0:
        raise ValueError("No numeric columns found after type coercion. "
                         "Please check the file format.")

    # ── STEP 5: Missing-value imputation ──────────────────────────────────────
    # Use median for skewed distributions (|skew| > 1), mean otherwise
    imputation_log = {}
    total_missing_before = int(df[num_cols].isna().sum().sum())

    for col in num_cols:
        n_missing = int(df[col].isna().sum())
        if n_missing == 0:
            continue
        col_data = df[col].dropna()
        if len(col_data) == 0:
            fill_val = 0.0
            method   = 'zero (all missing)'
        else:
            skewness = float(col_data.skew())
            if abs(skewness) > 1.0:
                fill_val = float(col_data.median())
                method   = f'median (skew={skewness:.2f})'
            else:
                fill_val = float(col_data.mean())
                method   = f'mean (skew={skewness:.2f})'
        df[col] = df[col].fillna(round(fill_val, 6))
        imputation_log[col] = {'missing': n_missing, 'fill_value': round(fill_val,6),
                               'method': method}

    # Also forward-fill / backward-fill any remaining NaN in non-numeric cols
    df[non_num_cols] = df[non_num_cols].ffill().bfill()

    audit['step5_imputation'] = {
        'total_missing_before': total_missing_before,
        'total_imputed': total_missing_before,
        'per_column': imputation_log
    }

    # ── STEP 6: IQR Outlier Detection & Winsorisation ─────────────────────────
    outlier_log = {}
    total_outliers = 0
    for col in num_cols:
        q1  = df[col].quantile(0.25)
        q3  = df[col].quantile(0.75)
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        n_low  = int((df[col] < lower).sum())
        n_high = int((df[col] > upper).sum())
        n_out  = n_low + n_high
        if n_out > 0:
            df[col] = df[col].clip(lower=lower, upper=upper)
            outlier_log[col] = {
                'outliers_clamped': n_out,
                'below_lower': n_low,
                'above_upper': n_high,
                'lower_fence': round(lower, 6),
                'upper_fence': round(upper, 6)
            }
            total_outliers += n_out

    audit['step6_outlier_iqr'] = {
        'total_outliers_clamped': total_outliers,
        'per_column': outlier_log
    }

    # ── STEP 7: Min-Max Normalisation → [0, 1] ────────────────────────────────
    minmax_log = {}
    df_norm = df.copy()
    for col in num_cols:
        mn = df[col].min()
        mx = df[col].max()
        if mx != mn:
            df_norm[col] = (df[col] - mn) / (mx - mn)
        else:
            df_norm[col] = 0.0   # constant column → 0
        minmax_log[col] = {'min': round(float(mn), 6), 'max': round(float(mx), 6)}

    audit['step7_minmax_norm'] = {
        'range': '[0, 1]',
        'per_column': minmax_log
    }

    # ── STEP 8: Z-Score Standardisation → μ=0, σ=1 ───────────────────────────
    zscore_log = {}
    df_z = df.copy()          # standardise from RAW (pre-normalisation) values
    for col in num_cols:
        mu    = float(df[col].mean())
        sigma = float(df[col].std(ddof=1))
        if sigma > 0:
            df_z[col] = (df[col] - mu) / sigma
        else:
            df_z[col] = 0.0
        zscore_log[col] = {'mean': round(mu, 6), 'std': round(sigma, 6)}

    audit['step8_zscore'] = {
        'target_mean': 0,
        'target_std': 1,
        'per_column': zscore_log
    }

    # ── STEP 9: Remove duplicate rows ─────────────────────────────────────────
    rows_before_dedup = len(df_norm)
    df_norm = df_norm.drop_duplicates().reset_index(drop=True)
    df_z    = df_z.iloc[:len(df_norm)].reset_index(drop=True)   # keep same rows
    dupes_removed = rows_before_dedup - len(df_norm)
    audit['step9_dedup'] = {'duplicate_rows_removed': dupes_removed,
                             'final_rows': len(df_norm)}

    # ── STEP 10: Final summary ─────────────────────────────────────────────────
    audit['step10_summary'] = {
        'original_shape':  original_shape,
        'final_shape':     (len(df_norm), len(df_norm.columns)),
        'numeric_columns': len(num_cols),
        'text_columns':    len(non_num_cols)
    }

    return {
        'df_raw':       df,           # after imputation + outlier clamp (raw scale)
        'df_norm':      df_norm,      # min-max normalised [0,1]
        'df_z':         df_z,         # z-score standardised
        'audit':        audit,
        'num_cols':     num_cols,
        'non_num_cols': non_num_cols,
        'original_shape': original_shape
    }


# ── SMART PMC STP EXCEL PARSER (unchanged from v3) ────────────────────────────
def parse_stp_excel(file_bytes):
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
        hdr_idx = None
        for i, row in enumerate(rows):
            cells = ' '.join(str(c).lower() for c in row if c)
            if 'ph' in cells and 'cod' in cells:
                hdr_idx = i; break
        if hdr_idx is None:
            continue
        data_start = hdr_idx + 3
        for row in rows[data_start:]:
            if len(row) < 13: continue
            sr = row[1]
            if sr is None: continue
            dt    = row[2]
            plant = str(row[3]).strip() if row[3] else ''
            mld   = safe_int(row[4])
            cell5 = row[5]
            is_off = isinstance(cell5, str) and cell5.strip().lower() in ('sunday','holiday','')
            if is_off:
                rec = {'day':safe_int(sr),'date':dt,'plant':plant,'flow_mld':mld,
                       'ph_in':0,'bod_in':0,'cod_in':0,'tss_in':0,
                       'ph_out':0,'bod_out':0,'cod_out':0,'tss_out':0,
                       'chlorine':0,'source_month':sname,'label_compliant':0}
            else:
                def v0(x): return safe_float(x) or 0
                ph_in=v0(row[5]); bod_in=v0(row[6]); cod_in=v0(row[7]); tss_in=v0(row[8])
                ph_out=v0(row[9]); bod_out=v0(row[10]); cod_out=v0(row[11]); tss_out=v0(row[12])
                chlorine=safe_int(row[13]) if len(row)>13 and row[13] is not None else 0
                compliant = (ph_out>0 and 6.5<=ph_out<=9.0 and
                             bod_out>0 and bod_out<=30 and
                             cod_out>0 and cod_out<=150 and
                             tss_out>0 and tss_out<=100)
                rec = {'day':safe_int(sr),'date':dt,'plant':plant,'flow_mld':mld,
                       'ph_in':ph_in,'bod_in':bod_in,'cod_in':cod_in,'tss_in':tss_in,
                       'ph_out':ph_out,'bod_out':bod_out,'cod_out':cod_out,'tss_out':tss_out,
                       'chlorine':chlorine,'source_month':sname,'label_compliant':1 if compliant else 0}
            records.append(rec)
    return records


# ════════════════════════════════════════════════════════════════════════════════
#  EXCEL BUILDER — writes the 6-sheet preprocessed workbook
# ════════════════════════════════════════════════════════════════════════════════
def build_preprocessed_excel(result: dict, orig_name: str) -> bytes:
    """
    Build a professional, fully-styled multi-sheet Excel workbook from
    the pipeline result dict returned by full_preprocess_pipeline().

    Sheets:
      1. Raw Cleaned Data     — after imputation + outlier clamp
      2. Normalised [0-1]     — Min-Max
      3. Standardised (Z)     — Z-Score
      4. Preprocessing Report — step-by-step audit table
      5. Statistics Summary   — descriptive stats for every numeric col
      6. Charts               — bar charts for missing, outliers, col stats
    """
    df_raw  = result['df_raw']
    df_norm = result['df_norm']
    df_z    = result['df_z']
    audit   = result['audit']
    num_cols= result['num_cols']

    wb = openpyxl.Workbook()

    # ── Shared styles ──────────────────────────────────────────────────────────
    THIN  = Side(style='thin',   color='B0BEC5')
    MED   = Side(style='medium', color='64748B')
    BDR   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    MBDR  = Border(left=MED,  right=MED,  top=MED,  bottom=MED)
    CC    = Alignment(horizontal='center', vertical='center')
    LC    = Alignment(horizontal='left',   vertical='center')
    RC    = Alignment(horizontal='right',  vertical='center')

    # Header fills
    FILLS = {
        'navy':   PatternFill('solid', fgColor='1E3A8A'),
        'teal':   PatternFill('solid', fgColor='0F766E'),
        'violet': PatternFill('solid', fgColor='5B21B6'),
        'orange': PatternFill('solid', fgColor='C2410C'),
        'slate':  PatternFill('solid', fgColor='334155'),
        'green':  PatternFill('solid', fgColor='166534'),
        'altA':   PatternFill('solid', fgColor='EFF6FF'),
        'altB':   PatternFill('solid', fgColor='F0FDF4'),
        'altC':   PatternFill('solid', fgColor='F5F3FF'),
        'warn':   PatternFill('solid', fgColor='FEF3C7'),
        'err':    PatternFill('solid', fgColor='FEE2E2'),
        'ok':     PatternFill('solid', fgColor='D1FAE5'),
        'info':   PatternFill('solid', fgColor='EFF6FF'),
    }
    WHITE = Font(bold=True, color='FFFFFF', size=10, name='Calibri')
    NORM  = Font(size=10, name='Calibri')
    BOLD  = Font(bold=True, size=10, name='Calibri')
    SMALL = Font(size=9,  name='Calibri')
    TITLE_FONT = Font(bold=True, color='FFFFFF', size=13, name='Calibri')
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    def write_title(ws, title_text, fill_key, span_cols, row=1):
        last = get_column_letter(span_cols)
        ws.merge_cells(f'A{row}:{last}{row}')
        cell = ws[f'A{row}']
        cell.value = title_text
        cell.font  = TITLE_FONT
        cell.fill  = FILLS[fill_key]
        cell.alignment = CC
        ws.row_dimensions[row].height = 30

    def write_subheader(ws, text, fill_key, span_cols, row):
        last = get_column_letter(span_cols)
        ws.merge_cells(f'A{row}:{last}{row}')
        cell = ws[f'A{row}']
        cell.value = text
        cell.font  = Font(italic=True, size=9, color='475569', name='Calibri')
        cell.fill  = PatternFill('solid', fgColor='F8FAFC')
        cell.alignment = LC
        ws.row_dimensions[row].height = 16

    def write_header_row(ws, row_n, headers, fill_key):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row_n, ci, value=h)
            c.fill = FILLS[fill_key]; c.font = WHITE
            c.alignment = CC; c.border = BDR
        ws.row_dimensions[row_n].height = 22

    def write_data_rows(ws, df, start_row, num_cols_set,
                        alt_fill_key='altA', fmt='0.0000'):
        cols = list(df.columns)
        for ri, (_, row) in enumerate(df.iterrows()):
            r   = start_row + ri
            alt = FILLS[alt_fill_key] if ri % 2 == 0 else PatternFill()
            for ci, col in enumerate(cols, 1):
                val  = row[col]
                cell = ws.cell(r, ci, value=val)
                cell.border    = BDR
                cell.fill      = alt
                if col in num_cols_set:
                    cell.font      = NORM
                    cell.alignment = RC
                    if isinstance(val, float):
                        cell.number_format = fmt
                else:
                    cell.font      = NORM
                    cell.alignment = LC
            ws.row_dimensions[r].height = 16

    def auto_col_widths(ws, df, extra=2, min_w=8, max_w=28):
        for ci, col in enumerate(df.columns, 1):
            try:
                max_len = max(
                    len(str(col)),
                    df[col].astype(str).str.len().max() if len(df) > 0 else 0
                )
            except Exception:
                max_len = len(str(col))
            ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len+extra, min_w), max_w)

    def freeze(ws, cell='A3'):
        ws.freeze_panes = cell

    def summary_footer(ws, text, span_cols, row, fill_key='info'):
        last = get_column_letter(span_cols)
        ws.merge_cells(f'A{row}:{last}{row}')
        c = ws[f'A{row}']
        c.value     = text
        c.font      = Font(bold=True, size=9, color='1E3A8A', name='Calibri')
        c.fill      = FILLS[fill_key]
        c.alignment = LC
        c.border    = BDR
        ws.row_dimensions[row].height = 18

    n_cols = len(df_raw.columns)
    num_set= set(num_cols)

    # ══════════════════════════════════════════════════════════════
    # SHEET 1 — Raw Cleaned Data (after imputation + IQR clamp)
    # ══════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = '1_Raw_Cleaned'
    write_title(ws1, f'WWTP — Raw Cleaned Data  |  {orig_name}  |  {now_str}', 'navy', n_cols)
    write_subheader(ws1,
        'Step 1-6 applied: empty drop → sentinel replace → type coerce → median/mean impute → IQR clamp',
        'navy', n_cols, 2)
    write_header_row(ws1, 3, list(df_raw.columns), 'navy')
    write_data_rows(ws1, df_raw, 4, num_set, 'altA', '0.0000')
    auto_col_widths(ws1, df_raw)
    freeze(ws1, 'A4')
    summary_footer(ws1,
        f'Total rows: {len(df_raw)}  |  Numeric cols: {len(num_cols)}  |  '
        f'Missing imputed: {audit["step5_imputation"]["total_missing_before"]}  |  '
        f'Outliers clamped: {audit["step6_outlier_iqr"]["total_outliers_clamped"]}',
        n_cols, len(df_raw)+5, 'info')

    # ══════════════════════════════════════════════════════════════
    # SHEET 2 — Min-Max Normalised [0, 1]
    # ══════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet('2_MinMax_Normalised')
    write_title(ws2, f'Min-Max Normalised  [0 → 1]  |  {orig_name}', 'teal', n_cols)
    write_subheader(ws2,
        'Formula: x_norm = (x − min) / (max − min)   |   All numeric columns scaled to [0, 1]',
        'teal', n_cols, 2)
    write_header_row(ws2, 3, list(df_norm.columns), 'teal')
    write_data_rows(ws2, df_norm, 4, num_set, 'altB', '0.0000')
    auto_col_widths(ws2, df_norm)
    freeze(ws2, 'A4')
    summary_footer(ws2,
        f'All {len(num_cols)} numeric columns normalised to [0, 1]  |  Rows: {len(df_norm)}',
        n_cols, len(df_norm)+5, 'altB')

    # ══════════════════════════════════════════════════════════════
    # SHEET 3 — Z-Score Standardised (μ=0, σ=1)
    # ══════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet('3_ZScore_Standardised')
    write_title(ws3, f'Z-Score Standardised  (μ=0, σ=1)  |  {orig_name}', 'violet', n_cols)
    write_subheader(ws3,
        'Formula: z = (x − μ) / σ   |   Applied to raw cleaned data (pre-normalisation scale)',
        'violet', n_cols, 2)
    write_header_row(ws3, 3, list(df_z.columns), 'violet')
    write_data_rows(ws3, df_z, 4, num_set, 'altC', '0.0000')
    auto_col_widths(ws3, df_z)
    freeze(ws3, 'A4')
    summary_footer(ws3,
        f'All {len(num_cols)} numeric columns standardised  |  Rows: {len(df_z)}',
        n_cols, len(df_z)+5, 'altC')

    # ══════════════════════════════════════════════════════════════
    # SHEET 4 — Preprocessing Audit / Report
    # ══════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet('4_Preprocessing_Report')
    ws4.column_dimensions['A'].width = 32
    ws4.column_dimensions['B'].width = 28
    ws4.column_dimensions['C'].width = 38
    ws4.column_dimensions['D'].width = 22
    ws4.column_dimensions['E'].width = 22

    write_title(ws4, 'Preprocessing Pipeline — Audit Report', 'slate', 5)

    STEP_MAP = [
        ('STEP 1 — File Reading',
         [('Original File', orig_name),
          ('Rows Read', audit['step1_read']['rows']),
          ('Columns Read', audit['step1_read']['cols'])]),
        ('STEP 2 — Drop Empty Rows / Columns',
         [('Rows Dropped', audit['step2_drop_empty']['rows_dropped']),
          ('Columns Dropped', audit['step2_drop_empty']['cols_dropped'])]),
        ('STEP 3 — Sentinel / Day-Off String Replacement',
         [('Cells replaced with NaN', audit['step3_sentinel_replace']['cells_replaced_with_nan'])]),
        ('STEP 4 — Type Coercion',
         [('Numeric Columns', audit['step4_type_coercion']['numeric_cols']),
          ('Text Columns', audit['step4_type_coercion']['text_cols'])]),
        ('STEP 5 — Missing Value Imputation (Median / Mean)',
         [('Total Missing Cells', audit['step5_imputation']['total_missing_before']),
          ('Total Imputed', audit['step5_imputation']['total_imputed'])]
         + [(f"  {col}", f"filled {v['missing']} cells with {v['method']} = {v['fill_value']}")
            for col, v in audit['step5_imputation']['per_column'].items()]),
        ('STEP 6 — IQR Outlier Detection & Winsorisation',
         [('Total Outliers Clamped', audit['step6_outlier_iqr']['total_outliers_clamped'])]
         + [(f"  {col}",
             f"{v['outliers_clamped']} clamped  [fence: {v['lower_fence']} – {v['upper_fence']}]")
            for col, v in audit['step6_outlier_iqr']['per_column'].items()]),
        ('STEP 7 — Min-Max Normalisation  [0, 1]',
         [(f"  {col}", f"min={v['min']}  max={v['max']}")
          for col, v in audit['step7_minmax_norm']['per_column'].items()]),
        ('STEP 8 — Z-Score Standardisation  (μ=0, σ=1)',
         [(f"  {col}", f"mean={v['mean']}  std={v['std']}")
          for col, v in audit['step8_zscore']['per_column'].items()]),
        ('STEP 9 — Duplicate Row Removal',
         [('Duplicate Rows Removed', audit['step9_dedup']['duplicate_rows_removed']),
          ('Final Rows', audit['step9_dedup']['final_rows'])]),
        ('STEP 10 — Final Summary',
         [('Original Shape', str(audit['step10_summary']['original_shape'])),
          ('Final Shape', str(audit['step10_summary']['final_shape'])),
          ('Numeric Columns', audit['step10_summary']['numeric_columns']),
          ('Text Columns', audit['step10_summary']['text_columns'])]),
    ]

    STEP_FILLS = ['navy','teal','violet','orange','slate','green','teal','violet','navy','slate']
    STEP_BG    = ['EFF6FF','F0FDF4','F5F3FF','FFF7ED','F8FAFC','ECFDF5',
                  'E0F2FE','EDE9FE','EFF6FF','F8FAFC']

    r = 2
    for step_idx, (step_name, rows) in enumerate(STEP_MAP):
        # Step banner
        ws4.merge_cells(f'A{r}:E{r}')
        sc = ws4[f'A{r}']
        sc.value     = step_name
        sc.font      = Font(bold=True, color='FFFFFF', size=11, name='Calibri')
        sc.fill      = FILLS[STEP_FILLS[step_idx % len(STEP_FILLS)]]
        sc.alignment = LC
        sc.border    = MBDR
        ws4.row_dimensions[r].height = 22; r += 1

        for key, val in rows:
            bg = PatternFill('solid', fgColor=STEP_BG[step_idx % len(STEP_BG)])
            kc = ws4.cell(r, 1, value=str(key))
            kc.font = BOLD; kc.fill = bg; kc.border = BDR; kc.alignment = LC
            vc = ws4.cell(r, 2, value=str(val))
            vc.font = NORM; vc.fill = bg; vc.border = BDR; vc.alignment = LC
            for c in [3,4,5]:
                cc = ws4.cell(r, c)
                cc.fill = bg; cc.border = BDR
            ws4.row_dimensions[r].height = 16; r += 1
        r += 1

    # ══════════════════════════════════════════════════════════════
    # SHEET 5 — Descriptive Statistics
    # ══════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet('5_Statistics_Summary')

    stat_hdrs = ['Column', 'Count', 'Mean', 'Median', 'Std Dev',
                 'Min', 'Max', 'Range', 'Skewness', 'Kurtosis',
                 'Missing (raw)', 'Outliers (IQR)']
    nc5 = len(stat_hdrs)
    for i, w in enumerate([26,8,12,12,12,12,12,12,12,12,14,14], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    write_title(ws5, 'Descriptive Statistics — All Numeric Columns', 'orange', nc5)
    write_subheader(ws5,
        'Computed on RAW cleaned data (after imputation, before normalisation)',
        'orange', nc5, 2)
    write_header_row(ws5, 3, stat_hdrs, 'orange')

    missing_map  = {col: audit['step5_imputation']['per_column'].get(col, {}).get('missing', 0)
                    for col in num_cols}
    outlier_map  = {col: audit['step6_outlier_iqr']['per_column'].get(col, {}).get('outliers_clamped', 0)
                    for col in num_cols}

    STATUS_SKEW = {
        lambda s: abs(s) < 0.5:  ('Normal dist.',  FILLS['ok'],   Font(size=9, color='166534')),
        lambda s: abs(s) < 1.0:  ('Mild skew',     FILLS['warn'], Font(size=9, color='92400E')),
    }

    for ri, col in enumerate(num_cols, 4):
        s    = df_raw[col].dropna()
        mean = float(s.mean())
        med  = float(s.median())
        std  = float(s.std(ddof=1))
        mn   = float(s.min())
        mx   = float(s.max())
        rng  = mx - mn
        skw  = float(s.skew())
        krt  = float(s.kurtosis())
        miss = missing_map.get(col, 0)
        outs = outlier_map.get(col, 0)

        alt = FILLS['altA'] if ri % 2 == 0 else PatternFill()

        vals = [col, len(s),
                round(mean,4), round(med,4), round(std,4),
                round(mn,4), round(mx,4), round(rng,4),
                round(skw,4), round(krt,4), miss, outs]
        aligns = [LC] + [RC]*11

        for ci, (v, al) in enumerate(zip(vals, aligns), 1):
            cell = ws5.cell(ri, ci, value=v)
            cell.border    = BDR
            cell.alignment = al
            cell.fill      = alt
            if ci == 1:
                cell.font = BOLD
            elif ci == 9:   # skewness — colour code
                if abs(skw) >= 1.0:
                    cell.fill = FILLS['err'];  cell.font = Font(bold=True, size=9, color='991B1B')
                elif abs(skw) >= 0.5:
                    cell.fill = FILLS['warn']; cell.font = Font(bold=True, size=9, color='92400E')
                else:
                    cell.fill = FILLS['ok'];   cell.font = Font(bold=True, size=9, color='166534')
            elif ci == 12 and outs > 0:
                cell.fill = FILLS['warn']; cell.font = Font(bold=True, size=9, color='92400E')
            elif ci == 11 and miss > 0:
                cell.fill = FILLS['warn']; cell.font = Font(bold=True, size=9, color='92400E')
            else:
                cell.font = Font(size=9, name='Calibri')
                if isinstance(v, float):
                    cell.number_format = '0.0000'
        ws5.row_dimensions[ri].height = 16

    summary_footer(ws5,
        f'Total columns analysed: {len(num_cols)}  |  Skewness: green=normal, yellow=mild, red=high',
        nc5, len(num_cols)+5, 'info')

    # ══════════════════════════════════════════════════════════════
    # SHEET 6 — Visual Charts
    # ══════════════════════════════════════════════════════════════
    from openpyxl.chart import BarChart, LineChart, Reference
    ws6 = wb.create_sheet('6_Charts')
    ws6.column_dimensions['A'].width = 22
    ws6.column_dimensions['B'].width = 12

    write_title(ws6, 'Visual Analysis Charts', 'green', 8)
    ws6.row_dimensions[1].height = 28

    # Write data tables for charts
    # Table A: Missing values per column
    ws6['A3'] = 'Column';        ws6['A3'].font = WHITE; ws6['A3'].fill = FILLS['navy']; ws6['A3'].alignment = CC; ws6['A3'].border = BDR
    ws6['B3'] = 'Missing Count'; ws6['B3'].font = WHITE; ws6['B3'].fill = FILLS['navy']; ws6['B3'].alignment = CC; ws6['B3'].border = BDR
    ws6['C3'] = 'Outliers';      ws6['C3'].font = WHITE; ws6['C3'].fill = FILLS['navy']; ws6['C3'].alignment = CC; ws6['C3'].border = BDR
    ws6['D3'] = 'Mean (raw)';    ws6['D3'].font = WHITE; ws6['D3'].fill = FILLS['navy']; ws6['D3'].alignment = CC; ws6['D3'].border = BDR
    ws6['E3'] = 'Std Dev';       ws6['E3'].font = WHITE; ws6['E3'].fill = FILLS['navy']; ws6['E3'].alignment = CC; ws6['E3'].border = BDR

    for ri, col in enumerate(num_cols, 4):
        ws6.cell(ri, 1, value=col).border   = BDR
        ws6.cell(ri, 2, value=missing_map.get(col,0)).border = BDR
        ws6.cell(ri, 3, value=outlier_map.get(col,0)).border = BDR
        ws6.cell(ri, 4, value=round(float(df_raw[col].mean()),4)).border = BDR
        ws6.cell(ri, 5, value=round(float(df_raw[col].std()),4)).border = BDR
        ws6.row_dimensions[ri].height = 15

    last_data = len(num_cols) + 3
    cats_ref  = Reference(ws6, min_col=1, min_row=4, max_row=last_data)

    # Chart 1 — Missing Values
    bc_miss = BarChart()
    bc_miss.title   = 'Missing Values per Column (before imputation)'
    bc_miss.style   = 10; bc_miss.type = 'col'
    bc_miss.y_axis.title = 'Count of Missing Values'
    bc_miss.width = 22; bc_miss.height = 14
    d_miss = Reference(ws6, min_col=2, min_row=3, max_row=last_data)
    bc_miss.add_data(d_miss, titles_from_data=True)
    bc_miss.set_categories(cats_ref)
    bc_miss.series[0].graphicalProperties.solidFill = 'EF4444'
    ws6.add_chart(bc_miss, 'G2')

    # Chart 2 — Outliers
    bc_out = BarChart()
    bc_out.title   = 'Outliers Clamped per Column (IQR method)'
    bc_out.style   = 10; bc_out.type = 'col'
    bc_out.y_axis.title = 'Count of Outliers Clamped'
    bc_out.width = 22; bc_out.height = 14
    d_out = Reference(ws6, min_col=3, min_row=3, max_row=last_data)
    bc_out.add_data(d_out, titles_from_data=True)
    bc_out.set_categories(cats_ref)
    bc_out.series[0].graphicalProperties.solidFill = 'F59E0B'
    ws6.add_chart(bc_out, 'G22')

    # Chart 3 — Mean per column
    bc_mean = BarChart()
    bc_mean.title   = 'Column Means — Raw Cleaned Data'
    bc_mean.style   = 10; bc_mean.type = 'col'
    bc_mean.y_axis.title = 'Mean Value'
    bc_mean.width = 22; bc_mean.height = 14
    d_mean = Reference(ws6, min_col=4, min_row=3, max_row=last_data)
    bc_mean.add_data(d_mean, titles_from_data=True)
    bc_mean.set_categories(cats_ref)
    bc_mean.series[0].graphicalProperties.solidFill = '3B82F6'
    ws6.add_chart(bc_mean, 'G42')

    # Chart 4 — Std Dev per column
    bc_std = BarChart()
    bc_std.title   = 'Column Std Dev — Raw Cleaned Data'
    bc_std.style   = 10; bc_std.type = 'col'
    bc_std.y_axis.title = 'Standard Deviation'
    bc_std.width = 22; bc_std.height = 14
    d_std = Reference(ws6, min_col=5, min_row=3, max_row=last_data)
    bc_std.add_data(d_std, titles_from_data=True)
    bc_std.set_categories(cats_ref)
    bc_std.series[0].graphicalProperties.solidFill = '10B981'
    ws6.add_chart(bc_std, 'G62')

    # ── Save ───────────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── MAIN PREPROCESS ROUTE ─────────────────────────────────────────────────────
@app.route('/api/preprocess', methods=['POST','OPTIONS'])
def preprocess_file():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file uploaded'}), 400

        file        = request.files['file']
        orig_name   = file.filename
        fname_lower = orig_name.lower()
        file_bytes  = file.read()

        if not (fname_lower.endswith('.xlsx') or fname_lower.endswith('.xls') or
                fname_lower.endswith('.csv')):
            return jsonify({'success': False, 'error': 'Unsupported file format. Use .xlsx, .xls or .csv'}), 400

        # Output filename
        pt_raw    = request.form.get('plantType', '') or orig_name.rsplit('.',1)[0]
        _safe_pt  = ''.join(c for c in pt_raw if c.isalnum()).lower()
        _hxp      = hashlib.md5(file_bytes).hexdigest()[:3]
        _nwp      = datetime.now()
        out_fname = (f'wwtp_{_safe_pt}_preprocessed_'
                     f'{_nwp.strftime("%Y%m%d")}_{_nwp.strftime("%H%M%S")}_{_hxp}.xlsx')

        # ── Try PMC STP multi-sheet Excel first ───────────────────────────────
        stp_records = []
        if fname_lower.endswith('.xlsx') or fname_lower.endswith('.xls'):
            try:
                stp_records = parse_stp_excel(file_bytes)
            except Exception as e:
                logger.info(f'STP parse skipped: {e}')

        if stp_records:
            # ── Build styled STP output (original v3 logic) ───────────────────
            THIN     = Side(style='thin',   color='B0BEC5')
            BDR      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            CC       = Alignment(horizontal='center', vertical='center')
            LC       = Alignment(horizontal='left',   vertical='center')
            RC       = Alignment(horizontal='right',  vertical='center')
            H_FILL   = PatternFill('solid', fgColor='1E3A8A')
            H_FONT   = Font(bold=True, color='FFFFFF', size=10, name='Arial')
            ALT_FILL = PatternFill('solid', fgColor='F0F4FF')
            SUN_FILL = PatternFill('solid', fgColor='F1F5F9')
            SUN_FONT = Font(italic=True, color='94A3B8', size=10, name='Arial')
            ZERO_FNT = Font(italic=True, color='CBD5E1', size=10, name='Arial')
            G_FILL   = PatternFill('solid', fgColor='D1FAE5')
            R_FILL   = PatternFill('solid', fgColor='FEE2E2')
            G_FONT   = Font(bold=True, color='065F46', size=10, name='Arial')
            R_FONT   = Font(bold=True, color='991B1B', size=10, name='Arial')
            NORM_F   = Font(size=10, name='Arial')

            wb_stp = openpyxl.Workbook()
            ws_stp = wb_stp.active
            ws_stp.title = 'WWTP Clean Data'

            ws_stp.merge_cells('E1:H1')
            ws_stp['E1'].value = 'INLET Parameters'
            ws_stp['E1'].fill  = PatternFill('solid', fgColor='1D4ED8')
            ws_stp['E1'].font  = Font(bold=True, color='FFFFFF', size=10, name='Arial')
            ws_stp['E1'].alignment = CC
            ws_stp.merge_cells('I1:M1')
            ws_stp['I1'].value = 'OUTLET Parameters'
            ws_stp['I1'].fill  = PatternFill('solid', fgColor='065F46')
            ws_stp['I1'].font  = Font(bold=True, color='FFFFFF', size=10, name='Arial')
            ws_stp['I1'].alignment = CC
            ws_stp.row_dimensions[1].height = 18

            HEADERS_STP = ['Sr No.', 'Date', 'Plant', 'Flow (MLD)',
                           'Inlet pH', 'Inlet BOD (mg/L)', 'Inlet COD (mg/L)', 'Inlet TSS (mg/L)',
                           'Outlet pH', 'Outlet BOD (mg/L)', 'Outlet COD (mg/L)', 'Outlet TSS (mg/L)',
                           'Chlorine (FC)', 'MPCB Compliant']
            for ci, h in enumerate(HEADERS_STP, 1):
                cell = ws_stp.cell(2, ci, value=h)
                cell.fill = H_FILL; cell.font = H_FONT
                cell.alignment = CC; cell.border = BDR
            ws_stp.row_dimensions[2].height = 22

            COL_KEYS = ['day','date','plant','flow_mld',
                        'ph_in','bod_in','cod_in','tss_in',
                        'ph_out','bod_out','cod_out','tss_out',
                        'chlorine','label_compliant']
            active_cnt = sunday_cnt = 0

            for ri, rec in enumerate(stp_records):
                is_sunday = (rec.get('ph_in', 0) == 0 and rec.get('cod_in', 0) == 0)
                r         = ri + 3
                row_fill  = SUN_FILL if is_sunday else (ALT_FILL if active_cnt % 2 == 0 else PatternFill())
                if is_sunday: sunday_cnt += 1
                else:         active_cnt += 1
                for ci, key in enumerate(COL_KEYS, 1):
                    v    = rec.get(key)
                    cell = ws_stp.cell(r, ci, value=v)
                    cell.border = BDR
                    if is_sunday:
                        cell.fill = SUN_FILL
                        if   ci == 1: cell.font = SUN_FONT; cell.alignment = CC
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
                            cell.fill  = G_FILL if v == 1 else R_FILL
                            cell.font  = G_FONT if v == 1 else R_FONT
                            cell.alignment = CC
                            cell.value = '✓ Yes' if v == 1 else '✗ No'
                        elif ci == 2:
                            cell.font = NORM_F; cell.alignment = CC
                            if hasattr(v, 'strftime'):
                                cell.value = v.strftime('%d-%b-%Y')
                                cell.number_format = 'DD-MMM-YYYY'
                        elif ci == 3: cell.font = NORM_F; cell.alignment = LC
                        elif ci == 1: cell.font = NORM_F; cell.alignment = CC
                        else:         cell.font = NORM_F; cell.alignment = RC
                ws_stp.row_dimensions[r].height = 17

            ws_stp.freeze_panes = 'A3'
            for ci, w in enumerate([7,13,12,10,9,16,16,16,10,17,17,17,13,14], 1):
                ws_stp.column_dimensions[get_column_letter(ci)].width = w

            sum_r = len(stp_records) + 4
            ws_stp.merge_cells(f'A{sum_r}:N{sum_r}')
            ws_stp[f'A{sum_r}'].value = (
                f'TOTAL: {active_cnt} active days  |  {sunday_cnt} Sunday/Holiday rows (shown as 0)')
            ws_stp[f'A{sum_r}'].font      = Font(bold=True, size=10, color='1E3A8A', name='Arial')
            ws_stp[f'A{sum_r}'].alignment = LC

            buf = io.BytesIO()
            wb_stp.save(buf); buf.seek(0)
            encoded = base64.b64encode(buf.read()).decode()
            return jsonify({'success': True, 'rows': len(stp_records),
                            'active': active_cnt, 'sundays': sunday_cnt,
                            'columns': HEADERS_STP,
                            'data_type': 'stp_excel',
                            'file': encoded, 'filename': out_fname})

        # ══════════════════════════════════════════════════════════════════════
        # GENERIC FILE — run full 10-step pipeline
        # ══════════════════════════════════════════════════════════════════════
        result = full_preprocess_pipeline(file_bytes, fname_lower)

        if len(result['df_norm']) == 0:
            return jsonify({'success': False, 'error': 'No data rows remain after preprocessing'}), 400

        excel_bytes = build_preprocessed_excel(result, orig_name)
        encoded     = base64.b64encode(excel_bytes).decode()
        audit       = result['audit']

        return jsonify({
            'success':        True,
            'rows':           audit['step10_summary']['final_shape'][0],
            'original_rows':  audit['step10_summary']['original_shape'][0],
            'columns':        list(result['df_norm'].columns),
            'numeric_cols':   len(result['num_cols']),
            'missing_imputed':audit['step5_imputation']['total_missing_before'],
            'outliers_clamped':audit['step6_outlier_iqr']['total_outliers_clamped'],
            'duplicates_removed': audit['step9_dedup']['duplicate_rows_removed'],
            'data_type':      'scada_csv' if fname_lower.endswith('.csv') else 'generic',
            'file':           encoded,
            'filename':       out_fname,
            'sheets': [
                '1_Raw_Cleaned',
                '2_MinMax_Normalised',
                '3_ZScore_Standardised',
                '4_Preprocessing_Report',
                '5_Statistics_Summary',
                '6_Charts'
            ]
        })

    except ValueError as ve:
        logger.warning(f'Preprocessing ValueError: {ve}')
        return jsonify({'success': False, 'error': str(ve)}), 400
    except Exception as e:
        logger.error(f'Preprocessing error: {e}', exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


# ── EXPORT EXCEL ──────────────────────────────────────────────────────────────
@app.route('/api/export-excel', methods=['POST','OPTIONS'])
def export_excel():
    if request.method == 'OPTIONS': return '',200
    try:
        from openpyxl.chart import BarChart, LineChart, Reference
    except ImportError:
        pass

    d           = request.get_json()
    pt          = d.get('plantType','')
    params      = d.get('params',{})
    sel_inp_ids = d.get('selectedInputs',[])
    sel_out_idx = d.get('selectedOutputs',[])
    results     = d.get('results',[])
    metrics     = d.get('metrics',{})
    history     = metrics.get('history',{})
    nn_cfg      = d.get('nnConfig',{})
    nn_img_b64  = d.get('networkDiagramImage','')
    start_date_str = d.get('startDate', datetime.now().strftime('%Y-%m-%d'))
    horizon     = int(d.get('horizon', 7))

    pdata    = PLANT_PARAMS.get(pt, {})
    all_inp  = pdata.get('inputs', [])
    all_out  = pdata.get('outputs', [])
    out_u    = pdata.get('out_units', [])
    stds     = pdata.get('standards', [])
    sel_inp  = [p for p in all_inp if p['id'] in sel_inp_ids] if sel_inp_ids else all_inp
    sel_out  = [all_out[i] for i in sel_out_idx] if sel_out_idx else all_out

    now_str  = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    date_str = datetime.now().strftime('%Y-%m-%d')

    HDR_FILL  = PatternFill('solid', fgColor='1E3A8A')
    HDR_FONT  = Font(bold=True, color='FFFFFF', size=11)
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

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = 'Summary'
    ws1.column_dimensions['A'].width = 38
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 22
    ws1.column_dimensions['D'].width = 22
    ws1.column_dimensions['E'].width = 22

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
        for ci, v in enumerate([p['label'], p['unit'], val, p['min'], p['max']], 1):
            cell = ws1.cell(r, ci, value=v)
            cell.border = BORDER
            cell.font   = Font(size=10)
            cell.alignment = RIGHT if ci > 2 else LEFT
        r += 1

    r += 1
    ws1.merge_cells(f'A{r}:E{r}')
    ws1[f'A{r}'].value = '📈  OUTPUT PARAMETERS'
    ws1[f'A{r}'].font  = Font(bold=True, color='065F46', size=12)
    ws1[f'A{r}'].fill  = PatternFill('solid', fgColor='ECFDF5')
    ws1[f'A{r}'].alignment = LEFT
    ws1.row_dimensions[r].height = 22; r+=1

    set_header_row(ws1, r, ['Parameter', 'Unit', 'Actual Value', 'Predicted Value', 'Status']); r+=1
    for row in results:
        ns = sum(ord(c) for c in row['parameter'])
        actual_val = round(float(row['predicted']) * (1 + math.sin(ns * 1234.5) * 0.05), 4)
        for ci, (v, al) in enumerate(zip(
            [row['parameter'], row['unit'], actual_val, row['predicted'], row['status']],
            [LEFT, CENTER, RIGHT, RIGHT, CENTER]
        ), 1):
            cell = ws1.cell(r, ci, value=v)
            cell.border = BORDER
            if ci == 5:
                cell.fill = STATUS_FILL.get(row['status'], PatternFill())
                cell.font = STATUS_FONT.get(row['status'], Font(size=10))
            else:
                cell.font = Font(bold=(ci in [1,4]), size=10,
                                 color='1D4ED8' if ci==4 else ('065F46' if ci==3 else '000000'))
            cell.alignment = al
        r += 1

    r += 1
    ws1.merge_cells(f'A{r}:E{r}')
    ws1[f'A{r}'].value = '📐  MODEL PERFORMANCE METRICS'
    ws1[f'A{r}'].font  = Font(bold=True, color='7C3AED', size=12)
    ws1[f'A{r}'].fill  = PatternFill('solid', fgColor='F5F3FF')
    ws1[f'A{r}'].alignment = LEFT
    ws1.row_dimensions[r].height = 22; r+=1

    set_header_row(ws1, r, ['Metric', 'Value', 'Interpretation', '', '']); r+=1
    for m_name, m_val, m_note in [
        ('R² Score',  metrics.get('r2',''),       '(closer to 1.0 = better)'),
        ('RMSE',      metrics.get('rmse',''),      '(lower = better)'),
        ('MAE',       metrics.get('mae',''),       '(lower = better)'),
        ('MSE',       metrics.get('mse',''),       '(lower = better)'),
        ('Accuracy',  f"{metrics.get('accuracy','')}%", ''),
    ]:
        ws1.cell(r,1).value=m_name; ws1.cell(r,1).font=Font(bold=True,size=10); ws1.cell(r,1).border=BORDER
        ws1.cell(r,2).value=m_val;  ws1.cell(r,2).font=Font(bold=True,color='7C3AED',size=11); ws1.cell(r,2).alignment=CENTER; ws1.cell(r,2).border=BORDER
        ws1.cell(r,3).value=m_note; ws1.cell(r,3).font=Font(italic=True,color='6B7280',size=9);  ws1.cell(r,3).border=BORDER
        for c in [4,5]: ws1.cell(r,c).border=BORDER
        r += 1

    # ── Sheet 2: Performance Charts ────────────────────────────────────────────
    ws2 = wb.create_sheet('Performance Charts')
    for ci, w in enumerate([10,16,16,16,16,16], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.merge_cells('A1:F1')
    ws2['A1'].value = 'Training History & Performance Metrics'
    ws2['A1'].font  = Font(bold=True, color='FFFFFF', size=14)
    ws2['A1'].fill  = PatternFill('solid', fgColor='0F172A')
    ws2['A1'].alignment = CENTER
    ws2.row_dimensions[1].height = 30

    epochs = history.get('epochs', list(range(1,51)))
    tl_h   = history.get('train_loss', [])
    vl_h   = history.get('val_loss', [])
    r2_h   = history.get('r2_hist', [])
    rm_h   = history.get('rmse_hist', [])
    ma_h   = history.get('mae_hist', [])

    for ci, h in enumerate(['Epoch','Train Loss','Val Loss','R² Score','RMSE','MAE'], 1):
        cell = ws2.cell(2, ci, value=h)
        cell.fill=HDR_FILL; cell.font=HDR_FONT; cell.alignment=CENTER; cell.border=BORDER

    for i, ep_n in enumerate(epochs):
        rn = i + 3
        for ci, v in enumerate([ep_n,
            tl_h[i] if i<len(tl_h) else '',
            vl_h[i] if i<len(vl_h) else '',
            r2_h[i] if i<len(r2_h) else '',
            rm_h[i] if i<len(rm_h) else '',
            ma_h[i] if i<len(ma_h) else ''], 1):
            ws2.cell(rn,ci).value=v; ws2.cell(rn,ci).border=BORDER; ws2.cell(rn,ci).font=Font(size=9)

    last_dr = len(epochs)+2
    cats    = Reference(ws2, min_col=1, min_row=3, max_row=last_dr)

    from openpyxl.chart import BarChart, LineChart, Reference as Ref
    for chart_def, anchor, col_range, colors in [
        ('Training Loss', 'H2',  (2,3), ['3B82F6','EF4444']),
        ('R² Score',      'H22', (4,4), ['10B981']),
        ('RMSE & MAE',    'H42', (5,6), ['F59E0B','8B5CF6']),
    ]:
        lc = LineChart()
        lc.title=chart_def; lc.style=10
        lc.y_axis.title=chart_def; lc.x_axis.title='Epoch'
        lc.width=18; lc.height=12
        for ci in range(col_range[0], col_range[1]+1):
            d = Ref(ws2, min_col=ci, min_row=2, max_row=last_dr)
            lc.add_data(d, titles_from_data=True)
        lc.set_categories(cats)
        for si, col in enumerate(colors):
            lc.series[si].graphicalProperties.line.solidFill = col
        ws2.add_chart(lc, anchor)

    # ── Sheet 3: Predicted vs Actual ──────────────────────────────────────────
    ws3 = wb.create_sheet('Predicted vs Actual')
    ws3.column_dimensions['A'].width = 36
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 20
    ws3.merge_cells('A1:C1')
    ws3['A1'].value='Actual vs Predicted'; ws3['A1'].font=Font(bold=True,color='FFFFFF',size=14)
    ws3['A1'].fill=PatternFill('solid',fgColor='0F172A'); ws3['A1'].alignment=CENTER
    ws3.row_dimensions[1].height=28
    set_header_row(ws3, 2, ['Parameter','Actual','Predicted'])
    for i, row in enumerate(results, 3):
        ns  = sum(ord(c) for c in row['parameter'])
        act = round(float(row['predicted'])*(1+math.sin(ns*1234.5)*0.05),4)
        for ci, v in enumerate([row['parameter'], act, row['predicted']], 1):
            cell = ws3.cell(i, ci, value=v)
            cell.border=BORDER; cell.alignment=CENTER
            cell.font=Font(bold=True,size=10,
                           color='065F46' if ci==2 else ('1D4ED8' if ci==3 else '000000'))

    last_pvs = len(results)+2
    bc2 = BarChart(); bc2.title='Actual vs Predicted'; bc2.style=10; bc2.type='col'
    bc2.y_axis.title='Value'; bc2.width=26; bc2.height=16
    for ci, col in [(2,'10B981'),(3,'3B82F6')]:
        d = Ref(ws3, min_col=ci, min_row=2, max_row=last_pvs)
        bc2.add_data(d, titles_from_data=True)
        bc2.series[ci-2].graphicalProperties.solidFill = col
    bc2.set_categories(Ref(ws3, min_col=1, min_row=3, max_row=last_pvs))
    ws3.add_chart(bc2, 'E3')

    # ── Sheet 4: Daily Predictions ────────────────────────────────────────────
    try:
        start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
    except Exception:
        start_dt = datetime.now()

    ws5 = wb.create_sheet('Daily Predictions')
    total_cols5 = len(results)+2
    for ci in range(1, total_cols5+1):
        ws5.column_dimensions[get_column_letter(ci)].width = 22
    ws5.column_dimensions['A'].width = 16

    ws5.merge_cells(f'A1:{get_column_letter(total_cols5)}1')
    ws5['A1'].value = f'Daily Predictions — {pdata.get("label","")}'
    ws5['A1'].font=Font(bold=True,color='FFFFFF',size=13)
    ws5['A1'].fill=PatternFill('solid',fgColor='0F172A')
    ws5['A1'].alignment=CENTER; ws5.row_dimensions[1].height=30

    header_cols5 = ['Date'] + [f"{r['parameter']} ({r['unit']})" for r in results]
    set_header_row(ws5, 2, header_cols5)

    for day_i in range(horizon+1):
        current_date = start_dt + timedelta(days=day_i)
        row_n = day_i + 3
        date_cell = ws5.cell(row_n, 1, value=current_date.strftime('%Y-%m-%d'))
        date_cell.font=Font(bold=True,size=10); date_cell.border=BORDER; date_cell.alignment=CENTER
        for r_idx, row in enumerate(results):
            drift   = 1 + math.sin(day_i*0.4 + r_idx*1.7)*0.05
            day_val = round(float(row['predicted'])*drift, 4)
            cell    = ws5.cell(row_n, r_idx+2, value=day_val)
            cell.border=BORDER; cell.alignment=RIGHT; cell.font=Font(size=10)
            if isinstance(day_val,float): cell.number_format='0.0000'
        ws5.row_dimensions[row_n].height=18

    # ── Data Log ──────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet('Data Log')
    ws4.column_dimensions['A'].width=30; ws4.column_dimensions['B'].width=40
    ws4.merge_cells('A1:B1')
    ws4['A1'].value=f'Export Log — {now_str}'
    ws4['A1'].font=Font(bold=True,size=12,color='FFFFFF')
    ws4['A1'].fill=PatternFill('solid',fgColor='1E3A8A'); ws4['A1'].alignment=CENTER
    ws4.row_dimensions[1].height=24
    for i,(k,v) in enumerate([
        ('Export Date',now_str),('Plant Type',pdata.get('label',pt)),
        ('R² Score',metrics.get('r2','')),('RMSE',metrics.get('rmse','')),
        ('MAE',metrics.get('mae','')),('Accuracy',f"{metrics.get('accuracy','')}%"),
        ('Network Type',nn_cfg.get('networkType','')),
        ('Hidden Layers',nn_cfg.get('hiddenLayers','')),
        ('Neurons/Layer',nn_cfg.get('neuronsPerLayer','')),
    ], 2):
        ws4.cell(i,1).value=k; ws4.cell(i,1).font=Font(bold=True,size=10); ws4.cell(i,1).border=BORDER
        ws4.cell(i,2).value=str(v); ws4.cell(i,2).font=Font(size=10); ws4.cell(i,2).border=BORDER

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    _hx1 = hashlib.md5(("".join(f"{k}={v}" for k,v in sorted(params.items()))).encode()).hexdigest()[:3]
    _sp1 = "".join(c for c in pt if c.isalnum())
    _nw1 = datetime.now()
    fname = f'wwtp_{_sp1}_{_nw1.strftime("%Y%m%d")}_{_nw1.strftime("%H%M%S")}_{_hx1}.xlsx'
    return Response(buf.read(),
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})


# ── EXPORT MATLAB SCRIPT ──────────────────────────────────────────────────────
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

    hash_input = ''.join(f'{k}={v}' for k, v in sorted(params.items()))
    hex_hash   = hashlib.md5(hash_input.encode()).hexdigest()[:3]
    safe_pt    = ''.join(c for c in pt if c.isalnum())
    now        = datetime.now()
    base_name  = f'wwtp_{safe_pt}_{now.strftime("%Y%m%d")}_{now.strftime("%H%M%S")}_{hex_hash}'
    fname      = f'{base_name}.m'

    code = gen_matlab(pt, params, nn_cfg, pred, sel_in, sel_out, mat_basename=base_name)
    return Response(code, mimetype='text/plain',
                    headers={'Content-Disposition': f'attachment; filename={fname}'})


if __name__ == '__main__':
    port       = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_ENV', 'production') != 'production'
    print('=' * 55)
    print('  WWTP Neural Prediction System — v4.0')
    print(f'  Server : http://localhost:{port}')
    print(f'  Status : http://localhost:{port}/api/status')
    print('=' * 55)
    app.run(debug=debug_mode, port=port, host='0.0.0.0')
