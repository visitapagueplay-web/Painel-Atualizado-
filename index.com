<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PaguePlay & Bookplay - Executive Intelligence</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root {
            --bg-gradient: linear-gradient(135deg, #02050e 0%, #0a1026 100%);
            --bg-card: rgba(13, 22, 47, 0.85);
            --bg-input: #0b132b;
            --text-primary: #ffffff;
            --text-secondary: #94a3b8;
            --accent-blue: #38bdf8;   
            --accent-green: #22c55e;  
            --accent-orange: #f97316;
            --accent-purple: #c084fc;
            --border-color: rgba(56, 189, 248, 0.15);
            --font-base: 'Segoe UI', system-ui, -apple-system, sans-serif;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            background: var(--bg-gradient);
            color: var(--text-primary);
            font-family: var(--font-base);
            padding: 30px;
            min-height: 100vh;
            -webkit-font-smoothing: antialiased;
        }

        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 20px;
        }

        .brand-container {
            display: flex;
            align-items: center;
            gap: 20px;
            background: rgba(255, 255, 255, 0.03);
            padding: 10px 20px;
            border-radius: 12px;
            border: 1px solid rgba(255, 255, 255, 0.05);
        }

        .brand-logo {
            height: 38px;
            object-fit: contain;
            filter: drop-shadow(0 0 8px rgba(56, 189, 248, 0.3));
        }

        .brand-divider {
            width: 1px;
            height: 30px;
            background-color: rgba(255, 255, 255, 0.15);
        }

        .title-area h1 {
            font-size: 24px;
            font-weight: 800;
            letter-spacing: -0.5px;
            text-align: right;
            background: linear-gradient(to right, #ffffff, #38bdf8);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .title-area p {
            font-size: 12px;
            color: var(--text-secondary);
            text-align: right;
        }

        /* --- SISTEMA DE ABAS --- */
        .tab-container {
            display: flex;
            gap: 15px;
            margin-bottom: 24px;
            justify-content: center;
        }
        .tab-btn {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-secondary);
            padding: 12px 30px;
            border-radius: 10px;
            font-size: 16px;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.3s;
        }
        .tab-btn.active {
            background: rgba(56, 189, 248, 0.15);
            color: #ffffff;
            border-color: var(--accent-blue);
            box-shadow: 0 0 15px rgba(56, 189, 248, 0.3);
        }

        /* --- CONTROLES DA ABA LIDERANÇA --- */
        .setores-bar {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
            flex-wrap: wrap;
            justify-content: center;
        }
        .setor-btn {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.1);
            color: #fff;
            padding: 10px 20px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.2s;
            text-transform: uppercase;
        }
        .setor-btn:hover {
            background: rgba(255,255,255,0.1);
        }
        .setor-btn.active {
            background: var(--accent-purple);
            border-color: var(--accent-purple);
            color: #000;
            box-shadow: 0 0 10px rgba(192, 132, 252, 0.4);
        }

        .lider-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 24px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.4);
            overflow: hidden;
        }
        .lider-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            padding-bottom: 15px;
            margin-bottom: 15px;
            flex-wrap: wrap;
            gap: 15px;
        }
        .lider-name {
            font-size: 20px;
            font-weight: 800;
            color: var(--accent-blue);
            text-transform: uppercase;
        }
        .lider-summary {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
        }
        .summary-item {
            background: rgba(0,0,0,0.3);
            padding: 10px 20px;
            border-radius: 8px;
            text-align: center;
            min-width: 120px;
        }
        .summary-label {
            font-size: 10px;
            color: var(--text-secondary);
            text-transform: uppercase;
            margin-bottom: 4px;
            font-weight: 700;
        }
        .summary-value {
            font-size: 22px;
            font-weight: 900;
        }

        /* --- ESTILOS ORIGINAIS --- */
        .action-container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 20px;
            margin-bottom: 24px;
            flex-wrap: wrap;
        }

        .control-bar {
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 12px 24px;
            display: flex;
            gap: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        }

        .control-group {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .control-group label {
            font-size: 10px;
            font-weight: 700;
            text-transform: uppercase;
            color: var(--text-secondary);
            letter-spacing: 0.6px;
        }

        .control-group input {
            background-color: var(--bg-input);
            border: 1px solid var(--border-color);
            color: var(--text-primary);
            padding: 6px 12px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 700;
            width: 90px;
            text-align: center;
            outline: none;
        }

        .import-actions {
            display: flex;
            gap: 12px;
        }

        .upload-wrapper {
            position: relative;
            overflow: hidden;
            display: inline-block;
        }

        .btn-premium {
            border: 1px solid var(--accent-blue);
            color: var(--accent-blue);
            background: rgba(56, 189, 248, 0.04);
            padding: 12px 18px;
            border-radius: 10px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .btn-premium:hover {
            background: var(--accent-blue);
            color: #000;
            box-shadow: 0 0 15px rgba(56, 189, 248, 0.35);
        }

        .btn-recorrente {
            border-color: var(--accent-purple);
            color: var(--accent-purple);
            background: rgba(192, 132, 252, 0.04);
        }

        .btn-recorrente:hover {
            background: var(--accent-purple);
            color: #000;
            box-shadow: 0 0 15px rgba(192, 132, 252, 0.35);
        }

        .upload-wrapper input[type=file] {
            position: absolute;
            left: 0; top: 0; opacity: 0; font-size: 100px; cursor: pointer;
        }

        .snapshot-container {
            display: flex;
            flex-direction: column;
            gap: 20px;
            margin-bottom: 32px;
        }

        .snapshot-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
        }

        .snapshot-row-regional {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
        }

        @media (max-width: 768px) {
            .snapshot-row-regional { grid-template-columns: 1fr; }
        }

        .kpi-card {
            background: linear-gradient(145deg, rgba(16, 29, 66, 0.9), rgba(7, 11, 26, 0.95));
            border: 1px solid var(--border-color);
            border-radius: 14px;
            padding: 22px;
            box-shadow: 0 12px 28px rgba(0,0,0,0.6);
            position: relative;
        }

        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; width: 4px; height: 100%;
            background: var(--accent-blue);
            border-radius: 14px 0 0 14px;
        }

        .kpi-card.kpi-total-geral::before { background: var(--accent-green); }
        .kpi-card.kpi-proporcional::before { background: var(--accent-orange); }
        .kpi-card.kpi-regional-mrl::before { background: var(--accent-purple); }
        .kpi-card.kpi-regional-brg::before { background: var(--accent-blue); }

        .kpi-label {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            color: var(--text-secondary);
            margin-bottom: 6px;
            font-weight: 700;
        }

        .kpi-value {
            font-size: 26px;
            font-weight: 800;
            letter-spacing: -0.5px;
            margin-bottom: 2px;
        }

        .kpi-meta-container {
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .btn-edit-meta {
            background: none;
            border: none;
            color: var(--accent-green);
            cursor: pointer;
            padding: 4px;
            border-radius: 6px;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .btn-edit-meta:hover {
            background: rgba(34, 197, 94, 0.15);
            color: #ffffff;
        }

        .kpi-value-huge {
            font-size: 36px;
            font-weight: 900;
            letter-spacing: -1px;
            text-shadow: 0 0 12px rgba(249, 115, 22, 0.2);
        }

        .kpi-subtext {
            font-size: 12px;
            color: var(--text-secondary);
            margin-top: 2px;
        }

        .regional-display {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            margin-bottom: 8px;
        }

        .kpi-value-regional {
            font-size: 28px;
            font-weight: 800;
            letter-spacing: -0.5px;
        }

        .kpi-pct-regional {
            font-size: 24px;
            font-weight: 800;
        }

        .kpi-subtext-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 12px;
            margin-top: 10px;
            padding-top: 10px;
            border-top: 1px solid rgba(255,255,255,0.06);
            font-size: 12px;
            color: var(--text-secondary);
        }

        .region-tag {
            position: absolute;
            top: 20px; right: 20px;
            font-size: 10px;
            font-weight: 800;
            text-transform: uppercase;
            padding: 3px 8px;
            border-radius: 5px;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(255,255,255,0.1);
        }

        .panel-header-title {
            font-size: 15px;
            font-weight: 700;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 6px;
        }

        .data-panel-card {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.5);
            margin-bottom: 32px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }

        th {
            background-color: rgba(56, 189, 248, 0.04);
            color: var(--text-secondary);
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
        }

        th.sortable {
            cursor: pointer;
            user-select: none;
        }

        th.sortable:hover {
            color: var(--text-primary);
            background-color: rgba(56, 189, 248, 0.08);
        }

        th.sortable .sort-icon {
            font-size: 9px;
            color: var(--accent-blue);
            margin-left: 2px;
        }

        td {
            padding: 14px 16px;
            font-size: 13px;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }

        tr:hover td {
            background-color: rgba(56, 189, 248, 0.02);
        }

        .td-valor-projetado-verde {
            color: var(--accent-green) !important;
            font-size: 14px;
            font-weight: 800 !important;
            text-shadow: 0 0 10px rgba(34, 197, 94, 0.15);
        }

        .highlight-pct {
            font-size: 16px;
            font-weight: 800;
            display: block;
        }

        .text-green { color: var(--accent-green); }
        .text-orange { color: var(--accent-orange); }
        .text-blue { color: var(--accent-blue); }

        .progress-bar-container {
            width: 100%;
            max-width: 120px;
            background: rgba(255, 255, 255, 0.05);
            height: 5px;
            border-radius: 3px;
            margin-top: 5px;
            overflow: hidden;
        }

        .progress-bar-fill {
            height: 100%;
            width: 0%;
            transition: width 0.3s;
        }

        /* --- RANKING DE LIDERANÇA --- */
        .ranking-header-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            flex-wrap: wrap;
            gap: 14px;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 16px 24px;
        }
        .ranking-sort-btn {
            background: rgba(255,255,255,0.05);
            border: 1px solid rgba(255,255,255,0.12);
            color: var(--text-secondary);
            padding: 8px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 12px;
            font-weight: 700;
            transition: all 0.2s;
        }
        .ranking-sort-btn.active {
            background: rgba(56, 189, 248, 0.15);
            border-color: var(--accent-blue);
            color: #fff;
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.25);
        }
        .ranking-table-container {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 10px 24px rgba(0, 0, 0, 0.5);
        }
        .ranking-table-header {
            display: grid;
            grid-template-columns: 40px 1fr 90px 80px 70px 80px;
            gap: 10px;
            align-items: center;
            background: linear-gradient(135deg, rgba(56, 189, 248, 0.15) 0%, rgba(34, 197, 94, 0.08) 100%);
            border-bottom: 2px solid rgba(56, 189, 248, 0.3);
            padding: 14px 16px;
            font-weight: 800;
            font-size: 11px;
            text-transform: uppercase;
            color: var(--accent-blue);
            letter-spacing: 0.8px;
            border-radius: 8px 8px 0 0;
            box-shadow: 0 2px 8px rgba(56, 189, 248, 0.1);
        }
        .ranking-table-row {
            display: grid;
            grid-template-columns: 40px 1fr 90px 80px 70px 80px;
            gap: 10px;
            align-items: center;
            padding: 13px 16px;
            border-bottom: 1px solid rgba(255,255,255,0.08);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            background: rgba(255,255,255,0.02);
        }
        .ranking-table-row:hover {
            background: rgba(56, 189, 248, 0.08);
            transform: translateX(4px);
            box-shadow: 0 2px 8px rgba(56, 189, 248, 0.15);
        }
        .ranking-table-row.rank-ouro {
            background: rgba(255, 215, 0, 0.06);
            border-left: 4px solid #ffd700;
        }
        .ranking-table-row.rank-prata {
            background: rgba(192, 192, 192, 0.05);
            border-left: 4px solid #c0c0c0;
        }
        .ranking-table-row.rank-bronze {
            background: rgba(205, 127, 50, 0.05);
            border-left: 4px solid #cd7f32;
        }
        .ranking-table-row.rank-bateu {
            background: rgba(34, 197, 94, 0.06);
            border-left: 4px solid #22c55e;
        }
        .ranking-pos {
            font-size: 18px;
            font-weight: 900;
            text-align: center;
        }
        .ranking-col-lider {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }
        .ranking-lider-nome {
            font-size: 12px;
            font-weight: 800;
            color: #fff;
            letter-spacing: 0.2px;
        }
        .ranking-setor-badge {
            display: inline-block;
            font-size: 8px;
            font-weight: 700;
            text-transform: uppercase;
            padding: 2px 6px;
            border-radius: 4px;
            background: rgba(56,189,248,0.12);
            border: 1px solid rgba(56,189,248,0.2);
            color: var(--accent-blue);
            width: fit-content;
        }
        .ranking-col-value {
            font-size: 13px;
            font-weight: 800;
            text-align: center;
        }
        .ranking-pct-box {
            display: flex;
            flex-direction: column;
            gap: 4px;
            align-items: center;
        }
        .ranking-pct-value {
            font-size: 14px;
            font-weight: 900;
        }
        .ranking-pct-bar {
            width: 100%;
            height: 4px;
            background: rgba(255,255,255,0.08);
            border-radius: 2px;
            overflow: hidden;
        }
        .ranking-pct-bar-fill {
            height: 100%;
            border-radius: 2px;
            transition: width 0.4s ease;
        }
        .ranking-badge-bateu {
            background: rgba(34,197,94,0.2);
            border: 1px solid rgba(34,197,94,0.4);
            color: #22c55e;
            font-size: 11px;
            font-weight: 800;
            padding: 4px 10px;
            border-radius: 6px;
            text-align: center;
            white-space: nowrap;
        }
        .ranking-status {
            text-align: center;
            font-size: 12px;
        }
        @media (max-width: 1200px) {
            .ranking-table-header,
            .ranking-table-row {
                grid-template-columns: 40px 120px 80px 70px 70px 90px;
                gap: 8px;
                font-size: 11px;
                padding: 10px 12px;
            }
        }
    </style>
</head>
<body>

    <header>
        <div class="brand-container">
            <img src="https://pagueplay.com.br/build/images/logo-pagueplay.svg" alt="PaguePlay Logo" class="brand-logo">
            <div class="brand-divider"></div>
            <img src="https://i.ibb.co/m5KJY4nB/Logo-do-Bookplay-Vertical-removebg-preview.png" alt="Bookplay Logo" class="brand-logo">
        </div>
        <div class="title-area">
            <h1>Painel de Operações</h1>
            <p>Sistema Integrado PaguePlay & Bookplay</p>
        </div>
    </header>

    <div class="tab-container">
        <button class="tab-btn active" id="btnTabGerencial" onclick="switchTab('gerencial')">Visão Gerencial</button>
        <button class="tab-btn" id="btnTabLideranca" onclick="switchTab('lideranca')">Visão Liderança</button>
        <button class="tab-btn" id="btnTabRanking" onclick="switchTab('ranking')">Ranking de Liderança</button>
    </div>

    <div id="tabGerencial">
        <div class="action-container">
            <section class="control-bar">
                <div class="control-group">
                    <label>Dias Corridos</label>
                    <input type="number" id="inputDiasPassados" value="7">
                </div>
                <div class="control-group">
                    <label>Dias Totais Mês</label>
                    <input type="number" id="inputDiasTotais" value="21">
                </div>
            </section>

            <div class="import-actions">
                <div class="upload-wrapper">
                    <button class="btn-premium">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
                        Importar PIX
                    </button>
                    <input type="file" id="uploadPix" accept=".xlsx, .xls, .csv">
                </div>
                <div class="upload-wrapper">
                    <button class="btn-premium btn-recorrente">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
                        Importar Recorrente
                    </button>
                    <input type="file" id="uploadRecorrente" accept=".xlsx, .xls, .csv">
                </div>
            </div>
        </div>

        <div class="snapshot-container">
            <div class="snapshot-row">
                <div class="kpi-card">
                    <div class="kpi-label">Valor Total Acumulado (Global)</div>
                    <div class="kpi-value" id="kpiValorGeral">R$ 0,00</div>
                    <div class="kpi-subtext">Soma Pix + Cartões</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Transações Processadas</div>
                    <div class="kpi-value" id="kpiVolTransacoes" style="color: var(--accent-blue);">0</div>
                    <div class="kpi-subtext">Volume integrado</div>
                </div>
                <div class="kpi-card kpi-total-geral">
                    <div class="kpi-label">Meta Unificada</div>
                    <div class="kpi-meta-container">
                        <div class="kpi-value" id="kpiMetaGlobal">R$ 3.000.000,00</div>
                        <button class="btn-edit-meta" onclick="alterarMetaUnificada()" title="Editar Meta Unificada">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 1 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                        </button>
                    </div>
                    <div class="kpi-subtext">Alvo estratégico</div>
                </div>
                <div class="kpi-card kpi-proporcional">
                    <div class="kpi-label">Atingimento do Ritmo (%)</div>
                    <div class="kpi-value kpi-value-huge" id="kpiPctGlobal" style="color: var(--accent-orange);">0,0%</div>
                    <div class="kpi-subtext" id="kpiDeveriaGeral">Deveria estar em: R$ 0,00</div>
                </div>
            </div>

            <div class="snapshot-row-regional">
                <div class="kpi-card kpi-regional-mrl">
                    <span class="region-tag" style="color: var(--accent-purple); border-color: var(--accent-purple);">Marília</span>
                    <div class="kpi-label">Desempenho Marília (Play 4 + 5)</div>
                    <div class="regional-display">
                        <div class="kpi-value-regional" id="regRealizadoMarilia">R$ 0,00</div>
                        <div class="kpi-pct-regional" id="regPctMarilia" style="color: var(--accent-green);">0,0%</div>
                    </div>
                    <div class="kpi-subtext-grid">
                        <div>Transações: <b id="regQtdMarilia" style="color:#fff;">0</b></div>
                        <div>Meta Mês: <b style="color:#fff;">R$ 500.000,00</b></div>
                        <div id="regDeveriaMarilia">Alvo Hoje: <b style="color:#fff;">R$ 0,00</b></div>
                    </div>
                </div>

                <div class="kpi-card kpi-regional-brg">
                    <span class="region-tag" style="color: var(--accent-blue); border-color: var(--accent-blue);">Birigui</span>
                    <div class="kpi-label">Desempenho Birigui (Demais Setores)</div>
                    <div class="regional-display">
                        <div class="kpi-value-regional" id="regRealizadoBirigui">R$ 0,00</div>
                        <div class="kpi-pct-regional" id="regPctBirigui" style="color: var(--accent-green);">0,0%</div>
                    </div>
                    <div class="kpi-subtext-grid">
                        <div>Transações: <b id="regQtdBirigui" style="color:#fff;">0</b></div>
                        <div>Meta Mês: <b style="color:#fff;">R$ 1.500.000,00</b></div>
                        <div id="regDeveriaBirigui">Alvo Hoje: <b style="color:#fff;">R$ 0,00</b></div>
                    </div>
                </div>
            </div>
        </div>

        <div class="panel-header-title" style="color: var(--accent-orange);">
            <span>Painel Master: Visão Geral Integrada</span>
        </div>
        <div class="data-panel-card">
            <table>
                <thead>
                    <tr>
                        <th class="sortable" onclick="triggerSort('global', 'nome')">Setor <span class="sort-icon">▲▼</span></th>
                        <th class="sortable" onclick="triggerSort('global', 'qtd')">Qtd Transações <span class="sort-icon">▲▼</span></th>
                        <th>Meta Setorial (R$)</th>
                        <th class="sortable" onclick="triggerSort('global', 'acumulado')">Valor Acumulado (Projeção Real) <span class="sort-icon">▲▼</span></th>
                        <th class="sortable" onclick="triggerSort('global', 'deveria')">Deveria Estar no Dia (R$) <span class="sort-icon">▲▼</span></th>
                        <th>Valor Pago</th>
                    </tr>
                </thead>
                <tbody id="tbodyGlobal"></tbody>
            </table>
        </div>

        <div class="panel-header-title" style="color: var(--accent-blue);">
            <span>Painel 02: Detalhamento Pix Automático</span>
        </div>
        <div class="data-panel-card">
            <table>
                <thead>
                    <tr>
                        <th class="sortable" onclick="triggerSort('pix', 'nome')">Setor <span class="sort-icon">▲▼</span></th>
                        <th>Quantidade Pix</th>
                        <th>Meta Setorial</th>
                        <th>Valor Acordo</th>
                        <th>Valor Pago</th>
                    </tr>
                </thead>
                <tbody id="tbodyPix"></tbody>
            </table>
        </div>

        <div class="panel-header-title" style="color: var(--accent-purple);">
            <span>Painel 03: Detalhamento Cartão Recorrente</span>
        </div>
        <div class="data-panel-card">
            <table>
                <thead>
                    <tr>
                        <th class="sortable" onclick="triggerSort('cartao', 'nome')">Setor <span class="sort-icon">▲▼</span></th>
                        <th>Transações</th>
                        <th class="sortable" onclick="triggerSort('cartao', 'bruto')">Valor Bruto Acumulado <span class="sort-icon">▲▼</span></th>
                        <th class="sortable" onclick="triggerSort('cartao', 'parcela')">Valor Parcela Acumulado (Recorrente) <span class="sort-icon">▲▼</span></th>
                    </tr>
                </thead>
                <tbody id="tbodyCartao"></tbody>
            </table>
        </div>
    </div>

    <div id="tabLideranca" style="display: none;">
        <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px; flex-wrap: wrap;">
            <div id="setoresBar"></div>
            <div style="margin-left: auto; display: flex; gap: 10px; align-items: center;">
                <input type="text" id="filtroOperador" placeholder="🔍 Filtrar operador..." style="padding: 8px 12px; background: var(--bg-input); border: 1px solid var(--border-color); border-radius: 8px; color: var(--text-primary); font-size: 13px;">
                <select id="filtroStatus" style="padding: 8px 12px; background: var(--bg-input); border: 1px solid var(--border-color); border-radius: 8px; color: var(--text-primary); font-size: 13px;">
                    <option value="">Todos os Status</option>
                    <option value="bateu">Bateu Meta ✓</option>
                    <option value="nao_bateu">Não Bateu</option>
                </select>
            </div>
        </div>
        <div id="liderancaContent"></div>
    </div>

    <div id="tabRanking" style="display: none;">
        <div class="ranking-header-bar">
            <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap;">
                <span style="font-size:18px; font-weight:800; color:var(--accent-blue);">🏆 Ranking de Liderança</span>
                <span style="font-size:13px; color:var(--text-secondary);">Quem está mais próximo de bater a meta?</span>
            </div>
            <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:12px; color:var(--text-secondary); font-weight:700; text-transform:uppercase;">Ordenar:</span>
                <button class="ranking-sort-btn active" id="btnSortDesc" onclick="setRankingSort('desc')">Maior % Primeiro ▼</button>
                <button class="ranking-sort-btn" id="btnSortAsc" onclick="setRankingSort('asc')">Menor % Primeiro ▲</button>
            </div>
        </div>
        <div id="rankingContent"></div>
    </div>

    <!-- RODAPÉ -->
    <footer style="margin-top: 60px; padding: 20px; text-align: center; border-top: 1px solid var(--border-color); color: var(--text-secondary); font-size: 12px;">
        <div>
            📅 <span id="dataAtual"></span> | 🕐 <span id="horaAtual"></span>
        </div>
        <div style="margin-top: 8px; font-size: 11px;">
            PaguePlay & Bookplay © 2024 - Sistema de Operações
        </div>
    </footer>

    <script>
        // --- ATUALIZAR DATA E HORA NO RODAPÉ ---
        function atualizarDataHora() {
            const agora = new Date();
            const opcoesDia = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
            const data = agora.toLocaleDateString('pt-BR', opcoesDia);
            const hora = agora.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            
            document.getElementById('dataAtual').innerText = data;
            document.getElementById('horaAtual').innerText = hora;
        }
        atualizarDataHora();
        setInterval(atualizarDataHora, 1000);

        // Lógica de Abas
        function switchTab(tab) {
            document.getElementById('tabGerencial').style.display = 'none';
            document.getElementById('tabLideranca').style.display = 'none';
            document.getElementById('tabRanking').style.display = 'none';
            document.getElementById('btnTabGerencial').classList.remove('active');
            document.getElementById('btnTabLideranca').classList.remove('active');
            document.getElementById('btnTabRanking').classList.remove('active');

            if (tab === 'gerencial') {
                document.getElementById('tabGerencial').style.display = 'block';
                document.getElementById('btnTabGerencial').classList.add('active');
            } else if (tab === 'lideranca') {
                document.getElementById('tabLideranca').style.display = 'block';
                document.getElementById('btnTabLideranca').classList.add('active');
                if (!window.liderancaIniciada) {
                    initLideranca();
                    window.liderancaIniciada = true;
                }
            } else if (tab === 'ranking') {
                document.getElementById('tabRanking').style.display = 'block';
                document.getElementById('btnTabRanking').classList.add('active');
                if (!window.liderancaIniciada) {
                    initLideranca();
                    window.liderancaIniciada = true;
                }
                renderRanking();
            }
        }

        // --- LÓGICA DA ABA LIDERANÇA (DADOS EMBUTIDOS) ---
        const rawOperadores = `Play1
JOZILAINE_LIMA
ariana_brito|rhaissa_sanches
Play1
RAIANA_REIS
ariana_brito|rhaissa_sanches
Play1
GEOVANNA_RAISSA
ariana_brito|rhaissa_sanches
Play1
CAMILLY_CARVALHO
ariana_brito|rhaissa_sanches
Play1
KIMBERLI_MARIN
ariana_brito|rhaissa_sanches
Play1
giovanni_fagundes
ariana_brito|rhaissa_sanches
Play1
KARINA_SILVA
ariana_brito|rhaissa_sanches
Play1
MAURO_SANTOS
ariana_brito|rhaissa_sanches
Play1
CAMILLY_PEREIRA
ariana_brito|rhaissa_sanches
Play1
LAIZA_DANIELE
ariana_brito|rhaissa_sanches
Play1
MONIQUE_LIMA
ariana_brito|rhaissa_sanches
Play1
CINTIA_DUARTEBRITO
ariana_brito|rhaissa_sanches
Play1
GABRIEL_MARTIN
ariana_brito|rhaissa_sanches
Play1
SHAKIRA_SANTOS
ariana_brito|rhaissa_sanches
Play1
ANA_DAMASCENO
ariana_brito|rhaissa_sanches
Play1
STEPHANIE_ANDRADE
ariana_brito|rhaissa_sanches
Play1
ALINE_RODRIGUES
ariana_brito|rhaissa_sanches
Play1
FERNANDA_MATA
ariana_brito|rhaissa_sanches
Play1
THIAGO_ALVES
ariana_brito|rhaissa_sanches
Play1
JOSE_SILVA
ariana_brito|rhaissa_sanches
Play1
EMILY_JARDIM
ariana_brito|rhaissa_sanches
Play1
THALIA_SOUZA
ariana_brito|rhaissa_sanches
Play1
LEANDRO_FERRAZ
mateus_castro|thais_venturin
Play1
CAMILA@RIBEIRO
mateus_castro|thais_venturin
Play1
ANDREINA_FRATUCCI
mateus_castro|thais_venturin
Play1
MARIA_SILVA
mateus_castro|thais_venturin
Play1
VITORIA_ARAUJO
mateus_castro|thais_venturin
Play1
ISABELLY_TELES
mateus_castro|thais_venturin
Play1
JULIANO_SILVA
mateus_castro|thais_venturin
Play1
MICAELA_SOUZA
mateus_castro|thais_venturin
Play1
JAQUELINE_BARBOSA
mateus_castro|thais_venturin
Play1
ROSINEIDE_MENDES
mateus_castro|thais_venturin
Play1
JULIA_CAETANO
mateus_castro|thais_venturin
Play1
CAROLINA_MURAI
mateus_castro|thais_venturin
Play1
ANDRIELE_SILVA
mateus_castro|thais_venturin
Play1
JULIANA_SPEREIRA
mateus_castro|thais_venturin
Play1
ANA_ZAMBELLI
mateus_castro|thais_venturin
Play1
MONIQUE_MELO
mateus_castro|thais_venturin
Play1
HEITOR_SANTOS
mateus_castro|thais_venturin
Play1
PRISCILA_ALVES
mateus_castro|thais_venturin
Play1
JULIA_BARBOZA
mateus_castro|thais_venturin
Play1
stella_nascimento
mateus_castro|thais_venturin
Play1
TAYNA_SOBRINHO
mateus_castro|thais_venturin
Play1
EMILLY_OLIVEIRA
mateus_castro|thais_venturin
Play1
ANA_DIAS
mateus_castro|thais_venturin
play2
NATIELE_NATAL
bianca_fernandes
play2
ISABELE_BARROS
bianca_fernandes
play2
POLIANA_SHUETER
bianca_fernandes
play2
MARIA_SEGURA
bianca_fernandes
play2
CLEITON_AFONSO
bianca_fernandes
play2
MARIHA_RODRIGUES
bianca_fernandes
play2
DANIELI_TREVIZAN
bianca_fernandes
play2
ROSANE_PIONA
bianca_fernandes
play2
AYSLAN_SILVA
bianca_fernandes
play2
LETICIA_S_GONCALVES
bianca_fernandes
play2
CAROLINA_GOBI
bianca_fernandes
play2
LOHAYNY_MAIA
bianca_fernandes
play2
laiza_santos
bianca_fernandes
play2
THAINA_COUTO
luan_barreto|luciano_costa
play2
LARISSA_VIEIRA
luan_barreto|luciano_costa
play2
RAYANE_ESTEVAN
luan_barreto|luciano_costa
play2
LETICIA_PERES
luan_barreto|luciano_costa
play2
MARIA_C_SILVA
luan_barreto|luciano_costa
play2
AMANDA_D_SILVA
luan_barreto|luciano_costa
play2
MIRIAN_BARBOSA
luan_barreto|luciano_costa
play2
AGATHA_MAZUCATO
luan_barreto|luciano_costa
play2
sabrina_lessa
luan_barreto|luciano_costa
play2
LETICIA_MATHIAS
luan_barreto|luciano_costa
play2
otavio_cestini
luan_barreto|luciano_costa
play2
THAMIRES_SOUZA
luan_barreto|luciano_costa
play2
DANIELA_REIS
luan_barreto|luciano_costa
play2
JHENIFER_ALMEIDA
luan_barreto|luciano_costa
play2
SHEILA_OLIVEIRA
luan_barreto|luciano_costa
play3
rayssa_ferreira
alisson_verissimo
play3
noemi_santos
alisson_verissimo
play3
andressa_aires
alisson_verissimo
play3
bruno_uemura
alisson_verissimo
play3
erika_eduarda
alisson_verissimo
play3
gabrieli_oliveira
alisson_verissimo
play3
luana_barrios
alisson_verissimo
play3
sthefanny_moraes
alisson_verissimo
play3
tainara_fernanda
alisson_verissimo
play3
thayna_izidro
alisson_verissimo
play3
beatriz_reis
ana_p_rodrigues
play3
luiz_silva
ana_p_rodrigues
play3
amanda_laisa
ana_p_rodrigues
play3
rebeca_machado
ana_p_rodrigues
play3
nicole_pacheco
ana_p_rodrigues
play3
tais_mantovanisantos
ana_p_rodrigues
play3
daniel_desouza
beatriz_bezerra
play3
isabeli_ferreira
beatriz_bezerra
play3
juliana_vitoria
beatriz_bezerra
play3
leticia_kerlia
beatriz_bezerra
play3
tuane_santos
beatriz_bezerra
play3
loami_ferraz
beatriz_bezerra
play3
giovanna_carvalho
beatriz_bezerra
play3
maria_e_teixeira
beatriz_bezerra
play3
mariaoliveiralima
beatriz_bezerra
play3
ana_caretta
sarah_picolin
play3
gabriela_gonçalves
sarah_picolin
play3
rodrigo_defalqui
sarah_picolin
play3
eduardo_naia
sarah_picolin
play3
gustavo_melo
sarah_picolin
play3
andre_brandao
sarah_picolin
Receptivo
GABRIEL_OLIVEIRA
bryan_queiroz
Receptivo
JOSE_VICTOR
bryan_queiroz
Receptivo
MARIANNE_FREITAS
bryan_queiroz
Receptivo
NAYARA_CRUZ
bryan_queiroz
Receptivo
EDUARDA_LORENZO
bryan_queiroz
Receptivo
paulo_silva
bryan_queiroz
Receptivo
KAUAN_TEIXEIRA
bryan_queiroz
Receptivo
NAYARA_MACEDO
luciana_machado
Receptivo
juliana_itala
luciana_machado
Receptivo
JENIFFER_OLIVEIRA
luciana_machado
Receptivo
heloisa_lima
luciana_machado
Receptivo
GABRIELY_ALVES
luciana_machado
Receptivo
maria_valeria
luciana_machado
Receptivo
RENATA_COSTA
luciana_machado
Receptivo
bianca_s_santos
luciana_machado
Receptivo
VIVIANE_ANTONIO
luciana_machado
Receptivo
eduardo_melo
matheus_costa
Receptivo
LARISSA_PEREIRAA
matheus_costa
Receptivo
LAYRA_CARINI
matheus_costa
Receptivo
FABIOLA_JAVAREZZI
matheus_costa
Receptivo
arielly_lima
matheus_costa
Receptivo
heloisa_camilo
matheus_costa
Receptivo
jeniffer_santos
matheus_costa
Receptivo
eriele_monteiro
matheus_costa
Receptivo
maria_mazziero
matheus_costa
Receptivo
thayla_silva
matheus_costa
Play6
FATIMA_TINARELI
douglas_faria
Play6
FABIO_GOIS
douglas_faria
Play6
ELAINE_SILVA
douglas_faria
Play6
SARAH_INOCENCIO
douglas_faria
Play6
CAMILA_COLANGELI
douglas_faria
Play6
VITORIA_SANTOS
douglas_faria
Play6
MATHEUS_MACEDO
douglas_faria
Play6
WILSON_JUNIOR
douglas_faria
Play6
stefany_souza
brunno_piccolo
Play6
GERCILENE_DIAS
brunno_piccolo
Play6
FRANCIELI_SILVA
brunno_piccolo
Play6
TIAGO_ALMADA
brunno_piccolo
Play6
ROSEMARY_DATRINO
brunno_piccolo
Play6
NAYARA_NUCCI
brunno_piccolo
Play6
JULIANA_BLASIOLI
brunno_piccolo
Play6
renansilva_goncalves
brunno_piccolo
Play6
JULIANA_S_MELLO
ellen_bruna
Play6
JULIANA_BELLINATI
ellen_bruna
Play6
VITORIA_GROSSO
ellen_bruna
Play6
JENNIFER_CUNHA
ellen_bruna
Play6
PAMELLA_MEDEIROS
ellen_bruna
Play6
QUEZIA_VITORELI
ellen_bruna
Play6
LUCILA_SILVA
ellen_bruna
Play6
EDNAN_CARLOS
ellen_bruna
Play6
TIAGO_HENRIQUE
ellen_bruna
Play6
rafael_pavanelli
ellen_bruna
Play4
VALQUIRIA_XAVIER
andreza_s_ribeiro
Play4
PRISCILA_KELLY
andreza_s_ribeiro
Play4
lilian_cavalcanti
andreza_s_ribeiro
Play4
LAURA_BARBOSA
andreza_s_ribeiro
Play4
caroline_loreti
andreza_s_ribeiro
Play4
SOPHIA_COSTA
tamires_valentin
Play4
MARIA_F_OLIVEIRA
tamires_valentin
Play4
ivanir_clarentinosou
tamires_valentin
Play4
BRUNA_C_OLIVEIRA
tamires_valentin
Play4
BARBARA_PONTOLIO
tamires_valentin
Play4
thaynaraoliveirasant
tauana_larissa
Play4
JANAINA_ARAUJO
tauana_larissa
Play4
ELISANDRA_RAQUEL
tauana_larissa
Play4
BRUNA_GARBI
tauana_larissa
Play4
BRUNA_TELES
tauana_larissa
play5
bruna_rafaela
camila_hitomi
play5
cintia_mengue
camila_hitomi
play5
gabrielly_barraca
camila_hitomi
play5
kethely_santos
camila_hitomi
play5
karina_colonez
camila_hitomi
play5
patricia_cruz
camila_hitomi
play5
sandy_rangel
camila_hitomi
play5
leticia_quinallia
camila_hitomi
play5
rebeca_camargo
layane_brito
play5
sirlei_stephanie
layane_brito
play5
claudineia_messias
layane_brito
play5
cleonice_goncalves
layane_brito
play5
fernanda_lopes
layane_brito
play5
julia_fernanda
layane_brito
play5
mariana_baleeiro
layane_brito
play5
milena_lima
layane_brito
play5
camila_isidoromoraes
layane_brito
play5
joao_anguita
layane_brito
play5
camila_rebelato
laysa_lopes
play5
caroline_valverde
laysa_lopes
play5
gloria_ferreira
laysa_lopes
play5
allana_barbosa
laysa_lopes
play5
juliana_s_oliveira
laysa_lopes
play5
silmara_silva
laysa_lopes
play5
yann_santos
laysa_lopes
play5
manuela_marini
laysa_lopes
play5
isabela_sobrinho
tauana_larissa
play5
rafaela_brumati
tauana_larissa
play5
maurinho
tauana_larissa
play5
eduarda_siviero
tauana_larissa
play5
tassiane_viana
tauana_larissa
play5
izadora_alexandre
tauana_larissa
PlayMix
GUSTAVO_NUNES
gabrielly_cheregatti
PlayMix
SOFIA_AMORIM
gabrielly_cheregatti
PlayMix
joao_garcia
gabrielly_cheregatti
PlayMix
MARIA_E_PEREIRA
gabrielly_cheregatti
PlayMix
MARIO_MANTOVANI
gabrielly_cheregatti
PlayMix
JOSIANI_JANUARIO
gabrielly_cheregatti
PlayMix
LUKAS_ALVES
gabrielly_cheregatti
PlayMix
PEDRO_ALVES
gabrielly_cheregatti`;

        function formatNome(nome) {
            return nome.replace(/_/g, ' ')
                       .replace(/@/g, ' ')
                       .toLowerCase()
                       .split(' ')
                       .map(word => word.charAt(0).toUpperCase() + word.slice(1))
                       .join(' ');
        }

        function normalizeNomeParaMatch(nome) {
            // Remove acentos, converte para lowercase, e padroniza com underscore
            return nome
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/\s+/g, '_')
                .replace(/@/g, '_')
                .replace(/_+/g, '_')
                .toLowerCase()
                .trim();
        }

        function parseOperadores(raw) {
            const lines = raw.trim().split('\n');
            const db = [];
            for(let i = 0; i < lines.length; i += 3) {
                if(!lines[i]) continue;
                let rawSetor = lines[i].trim();
                let setorNormalized = rawSetor.replace(/play\s*/i, 'Play ').trim();
                if(setorNormalized.toLowerCase() === 'playmix' || setorNormalized.toLowerCase() === 'play mix') {
                    setorNormalized = 'Play Mix';
                }

                let operador = lines[i+1].trim();
                let lider = lines[i+2].trim();

                // Tratamento especial para Laysa/Laisa
                lider = lider.replace(/la[iy]sa/gi, 'Laysa');

                let meta = 8;
                if(['Play 1', 'Play 2', 'Play 3'].includes(setorNormalized)) meta = 10;
                else if(['Play 4', 'Play 5', 'Play 6'].includes(setorNormalized)) meta = 6;

                db.push({
                    setor: setorNormalized,
                    operador: formatNome(operador),
                    operadorNormalized: normalizeNomeParaMatch(operador), // Adiciona versão normalizada
                    lider: formatNome(lider.replace(/\|/g, ' & ')),
                    meta: meta,
                    pix: 0,
                    cartao: 0,
                    transacoes: 0
                });
            }
            return db;
        }

        function initLideranca() {
            window.operadoresCompleto = parseOperadores(rawOperadores);
            
            const setoresArray = [...new Set(window.operadoresCompleto.map(o => o.setor))];
            
            const ordenarSetores = ["Play 1", "Play 2", "Play 3", "Play 4", "Play 5", "Play 6", "Play Mix", "Receptivo"];
            setoresArray.sort((a, b) => {
                let idxA = ordenarSetores.indexOf(a);
                let idxB = ordenarSetores.indexOf(b);
                if(idxA === -1) idxA = 99;
                if(idxB === -1) idxB = 99;
                return idxA - idxB;
            });

            const bar = document.getElementById('setoresBar');
            bar.innerHTML = setoresArray.map(s => 
                `<button class="setor-btn" onclick="selectSetorLideranca('${s}')" id="btn-setor-${s.replace(' ', '')}">${s}</button>`
            ).join('');

            if(setoresArray.length > 0) {
                selectSetorLideranca(setoresArray[0]);
            }
        }

        function selectSetorLideranca(setor) {
            document.querySelectorAll('.setor-btn').forEach(b => b.classList.remove('active'));
            const btn = document.getElementById('btn-setor-' + setor.replace(' ', ''));
            if(btn) btn.classList.add('active');

            // Resetar filtros
            document.getElementById('filtroOperador').value = '';
            document.getElementById('filtroStatus').value = '';

            const filtrados = window.operadoresCompleto.filter(o => o.setor === setor);
            const agrupado = {};
            filtrados.forEach(o => {
                if(!agrupado[o.lider]) agrupado[o.lider] = [];
                agrupado[o.lider].push(o);
            });

            const container = document.getElementById('liderancaContent');
            let html = '';

            // Painel de resumo do setor
            let totalSetorMeta = filtrados.reduce((acc, o) => acc + o.meta, 0);
            let totalSetorTrans = filtrados.reduce((acc, o) => acc + o.transacoes, 0);
            let gapSetorTotal = Math.max(0, totalSetorMeta - totalSetorTrans);
            let pctSetor = totalSetorMeta > 0 ? (totalSetorTrans / totalSetorMeta) * 100 : 0;

            html += `
            <div style="background: rgba(56, 189, 248, 0.08); border: 1px solid var(--border-color); border-radius: 12px; padding: 20px; margin-bottom: 24px; display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px;">
                <div style="text-align: center;">
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; margin-bottom: 8px;">Meta Total Setor</div>
                    <div style="font-size: 24px; font-weight: 900; color: var(--accent-blue);">${totalSetorMeta}</div>
                </div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; margin-bottom: 8px;">Realizado</div>
                    <div style="font-size: 24px; font-weight: 900; color: var(--accent-green);">${totalSetorTrans}</div>
                </div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; margin-bottom: 8px;">Distância</div>
                    <div style="font-size: 24px; font-weight: 900; color: var(--accent-orange);">${gapSetorTotal}</div>
                </div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; color: var(--text-secondary); text-transform: uppercase; margin-bottom: 8px;">% Realizado</div>
                    <div style="font-size: 24px; font-weight: 900; color: ${pctSetor >= 100 ? 'var(--accent-green)' : 'var(--accent-orange)'};">${pctSetor.toFixed(1)}%</div>
                </div>
            </div>
            `;

            for(let lider in agrupado) {
                let ops = agrupado[lider];
                // Ordenar operadores por total de transações (maior em primeiro)
                ops.sort((a, b) => b.transacoes - a.transacoes);
                
                let metaEquipe = ops.reduce((acc, curr) => acc + curr.meta, 0);
                let totalTrans = ops.reduce((acc, curr) => acc + curr.transacoes, 0);
                let gapEquipe = Math.max(0, metaEquipe - totalTrans);

                html += `
                <div class="lider-card">
                    <div class="lider-header">
                        <div class="lider-name">LÍDER: ${lider.toUpperCase()}</div>
                        <div class="lider-summary">
                            <div class="summary-item">
                                <div class="summary-label">Meta Equipe (Acordos)</div>
                                <div class="summary-value" style="color: var(--accent-blue);">${metaEquipe}</div>
                            </div>
                            <div class="summary-item">
                                <div class="summary-label">Realizado Total</div>
                                <div class="summary-value" style="color: var(--accent-green);">${totalTrans}</div>
                            </div>
                            <div class="summary-item">
                                <div class="summary-label">Distância / Gap</div>
                                <div class="summary-value" style="color: var(--accent-orange);">${gapEquipe}</div>
                            </div>
                        </div>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>Operador(a) <span style="font-size: 11px; color: var(--text-secondary);">⬇ Maior em 1º</span></th>
                                <th>Meta Individual</th>
                                <th>Pix Processados</th>
                                <th>Cartões Proc.</th>
                                <th>Total Transações</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${ops.map(o => {
                                let distancia = Math.max(0, o.meta - o.transacoes);
                                return `
                                <tr>
                                    <td style="font-weight: 600;">${o.operador}</td>
                                    <td style="color: var(--text-secondary);">${o.meta} acordos</td>
                                    <td>${o.pix}</td>
                                    <td>${o.cartao}</td>
                                    <td style="color: var(--accent-green); font-weight: bold; font-size: 15px;">${o.transacoes}</td>
                                    <td style="text-align: center; ${distancia > 0 ? 'color: var(--accent-orange);' : 'color: var(--accent-green);'} font-weight: bold;">${distancia > 0 ? 'Faltam ' + distancia : '✓ Batida'}</td>
                                </tr>
                                `;
                            }).join('')}
                        </tbody>
                    </table>
                </div>
                `;
            }
            container.innerHTML = html;
            
            // Adicionar event listeners para filtros
            const filtroOper = document.getElementById('filtroOperador');
            const filtroStat = document.getElementById('filtroStatus');
            
            if (filtroOper && filtroStat) {
                filtroOper.addEventListener('input', () => aplicarFiltrosLideranca(setor));
                filtroStat.addEventListener('change', () => aplicarFiltrosLideranca(setor));
            }
        }

        function aplicarFiltrosLideranca(setor) {
            const filtroOper = document.getElementById('filtroOperador').value.toLowerCase();
            const filtroStat = document.getElementById('filtroStatus').value;
            const linhas = document.querySelectorAll('table tbody tr');
            
            linhas.forEach(linha => {
                const operador = linha.querySelector('td:first-child')?.innerText.toLowerCase() || '';
                const status = linha.querySelector('td:last-child')?.innerText || '';
                const mostraOper = operador.includes(filtroOper);
                
                let mostraStat = true;
                if (filtroStat === 'bateu') {
                    mostraStat = status.includes('✓') || status.includes('Batida');
                } else if (filtroStat === 'nao_bateu') {
                    mostraStat = !status.includes('✓') && !status.includes('Batida');
                }
                
                linha.style.display = (mostraOper && mostraStat) ? '' : 'none';
            });
        }

        // --- RANKING DE LIDERANÇA ---
        window.rankingSortDir = 'desc';

        function setRankingSort(dir) {
            window.rankingSortDir = dir;
            document.getElementById('btnSortDesc').classList.toggle('active', dir === 'desc');
            document.getElementById('btnSortAsc').classList.toggle('active', dir === 'asc');
            renderRanking();
        }

        function renderRanking() {
            if (!window.operadoresCompleto || window.operadoresCompleto.length === 0) {
                if (!window.liderancaIniciada) {
                    initLideranca();
                    window.liderancaIniciada = true;
                }
            }

            // Agrupa todos os operadores por líder (sem filtrar setor)
            const agrupado = {};
            window.operadoresCompleto.forEach(o => {
                const key = o.lider;
                if (!agrupado[key]) {
                    agrupado[key] = { lider: o.lider, setores: new Set(), ops: [] };
                }
                agrupado[key].setores.add(o.setor);
                agrupado[key].ops.push(o);
            });

            // Calcula métricas por líder
            let lideresData = Object.values(agrupado).map(g => {
                const metaEquipe = g.ops.reduce((acc, o) => acc + o.meta, 0);
                const totalTrans = g.ops.reduce((acc, o) => acc + o.transacoes, 0);
                const pctRealizado = metaEquipe > 0 ? (totalTrans / metaEquipe) * 100 : 0;
                const gap = Math.max(0, metaEquipe - totalTrans);
                const setores = [...g.setores].join(', ');
                return { lider: g.lider, setores, metaEquipe, totalTrans, pctRealizado, gap, bateu: totalTrans >= metaEquipe };
            });

            // Ordena
            const dir = window.rankingSortDir === 'asc' ? 1 : -1;
            lideresData.sort((a, b) => (a.pctRealizado - b.pctRealizado) * dir);

            const medalhas = ['🥇', '🥈', '🥉'];
            const classPos = ['rank-ouro', 'rank-prata', 'rank-bronze'];

            let html = '<div class="ranking-table-container"><div class="ranking-table-header">';
            html += '<div style="text-align:center;">Pos</div>';
            html += '<div>Líder</div>';
            html += '<div style="text-align:center;">Meta Equipe</div>';
            html += '<div style="text-align:center;">Realizado</div>';
            html += '<div style="text-align:center;">% Realizado</div>';
            html += '<div style="text-align:center;">Status</div>';
            html += '</div>';

            lideresData.forEach((d, idx) => {
                const posLabel = window.rankingSortDir === 'desc'
                    ? (idx < 3 ? medalhas[idx] : `#${idx + 1}`)
                    : `#${idx + 1}`;
                const rankClass = d.bateu ? 'rank-bateu' : (window.rankingSortDir === 'desc' && idx < 3 ? classPos[idx] : 'rank-normal');

                const pctCapped = Math.min(d.pctRealizado, 100);
                const barColor = d.bateu ? '#22c55e' : (d.pctRealizado >= 75 ? '#38bdf8' : d.pctRealizado >= 40 ? '#f97316' : '#ef4444');
                const pctColor = d.bateu ? '#22c55e' : (d.pctRealizado >= 75 ? '#38bdf8' : d.pctRealizado >= 40 ? '#f97316' : '#ef4444');

                html += `
                <div class="ranking-table-row ${rankClass}">
                    <div class="ranking-pos">${posLabel}</div>
                    <div class="ranking-col-lider">
                        <div class="ranking-lider-nome">${d.lider.toUpperCase()}</div>
                        <div class="ranking-setor-badge">${d.setores}</div>
                    </div>
                    <div class="ranking-col-value" style="color: var(--accent-blue);">${d.metaEquipe}</div>
                    <div class="ranking-col-value" style="color: var(--accent-green);">${d.totalTrans}</div>
                    <div class="ranking-pct-box">
                        <div class="ranking-pct-value" style="color:${pctColor};">${d.pctRealizado.toFixed(1)}%</div>
                        <div class="ranking-pct-bar">
                            <div class="ranking-pct-bar-fill" style="width:${pctCapped}%; background:${barColor};"></div>
                        </div>
                    </div>
                    <div class="ranking-status">
                        ${d.bateu ? '<span class="ranking-badge-bateu">✓ META BATIDA</span>' : `<span style="color:var(--text-secondary); font-size:11px;">Faltam <b style="color:#fff;">${d.gap}</b></span>`}
                    </div>
                </div>`;
            });
            html += '</div>';
            document.getElementById('rankingContent').innerHTML = html;
        }

        // --- CÓDIGO ORIGINAL DA ABA GERENCIAL ---
        const metasOficiais = {
            "Play 1": 1457680.00, "Play 2": 1081750.00, "Play 3": 1502490.00,
            "Play 4": 73086.00, "Play 5": 337770.00, "Play 6": 208884.00,
            "Receptivo": 796168.00, "Play Mix": 291392.00
        };

        let META_GLOBAL_CONVERTIDA = 3000000.00; 
        const META_MARILIA = 500000.00;
        const META_BIRIGUI = 1500000.00;

        let DB_PIX = [
            { nome: "Play 1", qtd: 39, acordo: 162655.07, pago: 4614.83 },
            { nome: "Play 2", qtd: 27, acordo: 94920.86, pago: 4595.49 },
            { nome: "Play 3", qtd: 16, acordo: 65985.58, pago: 2106.54 },
            { nome: "Play 4", qtd: 0, acordo: 0.00, pago: 0.00 },
            { nome: "Play 5", qtd: 1, acordo: 2720.95, pago: 0.00 },
            { nome: "Play 6", qtd: 13, acordo: 22857.50, pago: 1127.44 },
            { nome: "Receptivo", qtd: 62, acordo: 216797.59, pago: 9602.22 },
            { nome: "Play Mix", qtd: 2, acordo: 9579.40, pago: 0.00 }
        ];

        let DB_CARTAO = [
            { nome: "Play 1", qtd: 366, bruto: 473169.16, parcela: 31544.61 },
            { nome: "Play 2", qtd: 140, bruto: 189244.33, parcela: 12616.28 },
            { nome: "Play 3", qtd: 75, bruto: 120937.69, parcela: 8062.51 },
            { nome: "Play 4", qtd: 18, bruto: 7631.00, parcela: 508.73 },
            { nome: "Play 5", qtd: 37, bruto: 33002.97, parcela: 2200.20 },
            { nome: "Play 6", qtd: 26, bruto: 47127.03, parcela: 3141.80 },
            { nome: "Receptivo", qtd: 248, bruto: 491105.21, parcela: 32740.35 },
            { nome: "Play Mix", qtd: 18, bruto: 17533.08, parcela: 1168.87 }
        ];

        let trackSort = { global: { col: null, asc: false }, pix: { col: null, asc: false }, cartao: { col: null, asc: false } };

        document.getElementById('inputDiasPassados').addEventListener('input', processingEngine);
        document.getElementById('inputDiasTotais').addEventListener('input', processingEngine);
        document.getElementById('uploadPix').addEventListener('change', (e) => triggerImport(e, 'pix'));
        document.getElementById('uploadRecorrente').addEventListener('change', (e) => triggerImport(e, 'cartao'));

        function formatFinancial(v) {
            return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(v);
        }

        function alterarMetaUnificada() {
            let novaMeta = prompt("Digite o novo valor para a Meta Unificada (Ex: 3000000):", META_GLOBAL_CONVERTIDA);
            if (novaMeta !== null) {
                let parsedMeta = parseFloat(novaMeta.replace(/[^0-9.]/g, ''));
                if (!isNaN(parsedMeta) && parsedMeta > 0) {
                    META_GLOBAL_CONVERTIDA = parsedMeta;
                    document.getElementById('kpiMetaGlobal').innerText = formatFinancial(META_GLOBAL_CONVERTIDA);
                    processingEngine();
                } else {
                    alert("Valor inválido inserido. Por favor, digite apenas números.");
                }
            }
        }

        function triggerSort(panel, col) {
            let state = trackSort[panel];
            state.asc = (state.col === col) ? !state.asc : false;
            state.col = col;

            const dp = parseFloat(document.getElementById('inputDiasPassados').value) || 1;
            const dt = parseFloat(document.getElementById('inputDiasTotais').value) || 21;

            let targetArr = panel === 'pix' ? DB_PIX : DB_CARTAO;

            if (panel === 'global') {
                DB_PIX.sort((a, b) => {
                    let cA = DB_CARTAO.find(i => i.nome === a.nome) || { qtd:0, bruto:0, parcela:0 };
                    let cB = DB_CARTAO.find(i => i.nome === b.nome) || { qtd:0, bruto:0, parcela:0 };
                    let vA, vB;

                    if (col === 'nome') return state.asc ? a.nome.localeCompare(b.nome) : b.nome.localeCompare(a.nome);
                    if (col === 'qtd') { vA = a.qtd + cA.qtd; vB = b.qtd + cB.qtd; }
                    if (col === 'acumulado') { vA = a.acordo + cA.bruto; vB = b.acordo + cB.bruto; }
                    if (col === 'deveria') {
                        vA = metasOficiais[a.nome] * (dp / dt);
                        vB = metasOficiais[b.nome] * (dp / dt);
                    }
                    return state.asc ? vA - vB : vB - vA;
                });
            } else {
                targetArr.sort((a, b) => {
                    if (col === 'nome') return state.asc ? a.nome.localeCompare(b.nome) : b.nome.localeCompare(a.nome);
                    let vA = a[col] || 0;
                    let vB = b[col] || 0;
                    return state.asc ? vA - vB : vB - vA;
                });
            }
            processingEngine();
        }

        function triggerImport(e, mode) {
            const file = e.target.files[0];
            if (!file) return;
            
            // Garante que liderança foi inicializada
            if (!window.liderancaIniciada) {
                initLideranca();
                window.liderancaIniciada = true;
            }
            
            const reader = new FileReader();
            reader.onload = function(evt) {
                const workbook = XLSX.read(evt.target.result, { type: 'binary' });
                const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
                
                let mapping = {
                    "Play 1": {q:0, v:0, p:0, bruto:0, parcela:0}, "Play 2": {q:0, v:0, p:0, bruto:0, parcela:0}, "Play 3": {q:0, v:0, p:0, bruto:0, parcela:0},
                    "Play 4": {q:0, v:0, p:0, bruto:0, parcela:0}, "Play 5": {q:0, v:0, p:0, bruto:0, parcela:0}, "Play 6": {q:0, v:0, p:0, bruto:0, parcela:0},
                    "Receptivo": {q:0, v:0, p:0, bruto:0, parcela:0}, "Play Mix": {q:0, v:0, p:0, bruto:0, parcela:0}
                };

                let headerIdx = 0;
                for (let i = 0; i < Math.min(rawRows.length, 10); i++) {
                    let rowStr = (rawRows[i] || []).join(' ').toLowerCase();
                    if ((rowStr.includes('setor') || rowStr.includes('grupo') || rowStr.includes('origem')) && rowStr.includes('valor')) {
                        headerIdx = i;
                        break;
                    }
                }

                const headers = rawRows[headerIdx] || [];
                let colGrupo = -1;
                let colsValor = [];
                let colPago = -1;
                let colOper = -1; // Coluna de operador/nome
                
                headers.forEach((h, idx) => {
                    if (typeof h === 'string') {
                        let hl = h.toLowerCase().trim();
                        if (hl === 'setor' || hl === 'grupo' || hl === 'equipe' || hl === 'origem' || hl === 'categoria' || hl === 'local') colGrupo = idx;
                        if (hl === 'valor acordo' || hl === 'valor bruto' || hl === 'valor parcela' || hl === 'valor') colsValor.push(idx);
                        if (hl === 'valor pago') colPago = idx;
                        // Detecta coluna de operador/nome
                        if (hl.includes('operador') || hl.includes('login') || hl.includes('usuario') || hl.includes('usuário') || hl.includes('nome do operador')) {
                            colOper = idx;
                        }
                    }
                });

                // Mapeamento setorial + coleta de nomes de operadores
                const operadoresNoArquivo = {};

                for (let i = headerIdx + 1; i < rawRows.length; i++) {
                    let r = rawRows[i];
                    if (!r || r.length === 0) continue;
                    
                    let label = '';
                    if (colGrupo !== -1 && typeof r[colGrupo] === 'string') {
                        label = r[colGrupo].toUpperCase();
                    } else {
                        for (let cell of r) {
                            if (typeof cell === 'string') {
                                let s = cell.toUpperCase();
                                // Procura por padrões conhecidos
                                if (s.includes('COB PLAY') || s.includes('COB RECEPTIVO') || s.includes('MARILIA') || s.includes('MARÍLIA') || 
                                    s.includes('PLAYMIX') || s.includes('PLAY MIX') || s.includes('EM DIA') || s.includes('RECEPTIVO')) {
                                    label = s;
                                    break;
                                }
                            }
                        }
                    }

                    if (!label) continue;

                    let k = null;
                    // Detectar padrões específicos dos arquivos
                    if (label.includes('COB PLAY 1') || label.includes('PLAY 1')) k = 'Play 1';
                    else if (label.includes('COB PLAY 2') || label.includes('PLAY 2')) k = 'Play 2';
                    else if (label.includes('COB PLAY 3') || label.includes('PLAY 3') || label.includes('JAQUELINI')) k = 'Play 3';
                    else if (label.includes('COB PLAY 6') || label.includes('PLAY 6')) k = 'Play 6';
                    else if (label.includes('MARILIA') || label.includes('MARÍLIA')) {
                        if (label.includes('PLAY 5') || label.includes('5')) k = 'Play 5';
                        else if (label.includes('PLAY 4') || label.includes('4')) k = 'Play 4';
                        else k = 'Play 5'; // padrão Marília
                    }
                    else if (label.includes('COB RECEPTIVO') || label.includes('COB PLAY - EM DIA') || 
                             label.includes('EM DIA') || label.includes('RECEPTIVO') || label.includes('RECEPTIVA') || 
                             label.includes('TELEMARKETING') || label.includes('INBOUND')) k = 'Receptivo';
                    else if (label.includes('MIX') || label.includes('PLAYMIX') || label.includes('PLAY MIX')) k = 'Play Mix';

                    if (!k) continue;

                    // Coleta nome do operador se disponível
                    if (colOper !== -1 && r[colOper]) {
                        let nomeOper = (r[colOper] || '').toString().trim();
                        if (nomeOper) {
                            let nomeNormalizado = normalizeNomeParaMatch(nomeOper);
                            operadoresNoArquivo[nomeNormalizado] = (operadoresNoArquivo[nomeNormalizado] || 0) + 1;
                        }
                    }

                    if (mode === 'pix') {
                        let valAcordo = parseFloat(r[colsValor[0]]) || 0;
                        let valPago = colPago !== -1 ? (parseFloat(r[colPago]) || 0) : valAcordo;
                        if (valAcordo > 0) {
                            mapping[k].q++;
                            mapping[k].v += valAcordo;
                            mapping[k].p += valPago;
                        }
                    } else {
                        let idxBruto = headers.findIndex(h => typeof h === 'string' && h.toLowerCase().trim() === 'valor bruto');
                        let idxParcela = headers.findIndex((h, idx) => typeof h === 'string' && h.toLowerCase().trim() === 'valor' && idx > idxBruto);
                        if (idxParcela === -1) {
                            idxParcela = headers.findIndex(h => typeof h === 'string' && h.toLowerCase().trim() === 'valor');
                        }

                        let brutoVal = idxBruto !== -1 ? (parseFloat(r[idxBruto]) || 0) : 0;
                        let parcelaVal = idxParcela !== -1 ? (parseFloat(r[idxParcela]) || 0) : 0;

                        if (brutoVal > 0 || parcelaVal > 0) {
                            mapping[k].q++;
                            mapping[k].bruto += brutoVal;
                            mapping[k].parcela += parcelaVal;
                        }
                    }
                }

                if (mode === 'pix') {
                    DB_PIX = Object.keys(mapping).map(key => ({ nome: key, qtd: mapping[key].q, acordo: mapping[key].v, pago: mapping[key].p }));
                } else {
                    DB_CARTAO = Object.keys(mapping).map(key => ({ nome: key, qtd: mapping[key].q, bruto: mapping[key].bruto, parcela: mapping[key].parcela }));
                }

                // Atualiza transações nos operadores usando normalização para match
                if (window.operadoresCompleto && Object.keys(operadoresNoArquivo).length > 0) {
                    window.operadoresCompleto.forEach(o => {
                        const nomeNorm = o.operadorNormalized;
                        const contagem = operadoresNoArquivo[nomeNorm] || 0;
                        
                        if (mode === 'pix') { 
                            o.pix = contagem; 
                        } else { 
                            o.cartao = contagem; 
                        }
                        o.transacoes = o.pix + o.cartao;
                    });

                    // Atualiza visão liderança se estiver visível
                    const tabLid = document.getElementById('tabLideranca');
                    if (tabLid && tabLid.style.display !== 'none') {
                        const setorAtivo = document.querySelector('.setor-btn.active');
                        if (setorAtivo) selectSetorLideranca(setorAtivo.textContent.trim());
                    }
                    // Atualiza ranking SEMPRE
                    renderRanking();
                }

                processingEngine();
            };
            reader.readAsBinaryString(file);
        }

        function processingEngine() {
            const dp = parseFloat(document.getElementById('inputDiasPassados').value) || 1;
            const dt = parseFloat(document.getElementById('inputDiasTotais').value) || 21;

            let outGlobal = '', outPix = '', outCartao = '';
            let totalRealizadoGeral = 0, totalTransacoesGeral = 0;
            
            let realizadoMarilia = 0, qtdMarilia = 0;
            let realizadoBirigui = 0, qtdBirigui = 0;

            DB_PIX.forEach(p => {
                let c = DB_CARTAO.find(i => i.nome === p.nome) || { qtd: 0, bruto: 0, parcela: 0 };
                let sumQtd = p.qtd + c.qtd;
                let sumValue = p.acordo + c.bruto; 
                let metaSetorial = metasOficiais[p.nome] || 0;

                totalRealizadoGeral += sumValue;
                totalTransacoesGeral += sumQtd;

                if (p.nome === "Play 4" || p.nome === "Play 5") {
                    realizadoMarilia += sumValue;
                    qtdMarilia += sumQtd;
                } else {
                    realizadoBirigui += sumValue;
                    qtdBirigui += sumQtd;
                }

                let deveriaEstarNoDia = metaSetorial * (dp / dt);
                let pctDeveria = deveriaEstarNoDia > 0 ? (sumValue / deveriaEstarNoDia) * 100 : 0;

                let color = 'text-orange';
                if (pctDeveria >= 100) color = 'text-green';
                else if (pctDeveria >= 80) color = 'text-blue';

                outGlobal += `
                    <tr>
                        <td style="font-weight:700;">${p.nome}</td>
                        <td>${sumQtd}</td>
                        <td style="color:var(--text-secondary); font-weight:600;">${formatFinancial(metaSetorial)}</td>
                        <td class="td-valor-projetado-verde">${formatFinancial(sumValue)}</td>
                        <td>
                            <span class="highlight-pct ${color}">${pctDeveria.toFixed(1)}%</span>
                            <span style="font-size:11px; color:var(--text-secondary);">Alvo do Dia: ${formatFinancial(deveriaEstarNoDia)}</span>
                            <div class="progress-bar-container">
                                <div class="progress-bar-fill" style="width: ${Math.min(pctDeveria, 100)}%; background-color: var(--${color === 'text-green' ? 'accent-green' : color === 'text-blue' ? 'accent-blue' : 'accent-orange'})"></div>
                            </div>
                        </td>
                        <td style="font-weight:600; color:var(--accent-green);">${formatFinancial(p.pago)}</td>
                    </tr>
                `;
            });

            DB_PIX.forEach(p => {
                outPix += `<tr><td style="font-weight:600;">${p.nome}</td><td>${p.qtd}</td><td>${formatFinancial(metasOficiais[p.nome])}</td><td>${formatFinancial(p.acordo)}</td><td>${formatFinancial(p.pago)}</td></tr>`;
            });

            DB_CARTAO.forEach(c => {
                outCartao += `<tr><td style="font-weight:600;">${c.nome}</td><td>${c.qtd}</td><td>${formatFinancial(c.bruto)}</td><td style="color:var(--accent-purple); font-weight:600;">${formatFinancial(c.parcela)}</td></tr>`;
            });

            document.getElementById('tbodyGlobal').innerHTML = outGlobal;
            document.getElementById('tbodyPix').innerHTML = outPix;
            document.getElementById('tbodyCartao').innerHTML = outCartao;

            let globalDeveria = META_GLOBAL_CONVERTIDA * (dp / dt);
            let globalPct = globalDeveria > 0 ? (totalRealizadoGeral / globalDeveria) * 100 : 0;

            document.getElementById('kpiValorGeral').innerText = formatFinancial(totalRealizadoGeral);
            document.getElementById('kpiVolTransacoes').innerText = totalTransacoesGeral.toLocaleString('pt-BR');
            document.getElementById('kpiDeveriaGeral').innerText = `Deveria estar em: ${formatFinancial(globalDeveria)}`;
            document.getElementById('kpiPctGlobal').innerText = globalPct.toFixed(1) + '%';

            let deveriaMarilia = META_MARILIA * (dp / dt);
            let pctMarilia = deveriaMarilia > 0 ? (realizadoMarilia / deveriaMarilia) * 100 : 0;
            document.getElementById('regRealizadoMarilia').innerText = formatFinancial(realizadoMarilia);
            document.getElementById('regQtdMarilia').innerText = qtdMarilia.toLocaleString('pt-BR');
            document.getElementById('regPctMarilia').innerText = pctMarilia.toFixed(1) + '%';
            document.getElementById('regDeveriaMarilia').innerHTML = `Alvo Hoje: <b style="color:#fff;">${formatFinancial(deveriaMarilia)}</b>`;

            let deveriaBirigui = META_BIRIGUI * (dp / dt);
            let pctBirigui = deveriaBirigui > 0 ? (realizadoBirigui / deveriaBirigui) * 100 : 0;
            document.getElementById('regRealizadoBirigui').innerText = formatFinancial(realizadoBirigui);
            document.getElementById('regQtdBirigui').innerText = qtdBirigui.toLocaleString('pt-BR');
            document.getElementById('regPctBirigui').innerText = pctBirigui.toFixed(1) + '%';
            document.getElementById('regDeveriaBirigui').innerHTML = `Alvo Hoje: <b style="color:#fff;">${formatFinancial(deveriaBirigui)}</b>`;
        }

        processingEngine();
    </script>
</body>
</html>
