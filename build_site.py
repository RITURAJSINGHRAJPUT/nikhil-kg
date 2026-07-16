import json
import os

def build():
    # Read the parsed machines data
    with open('machines.json', 'r', encoding='utf-8') as f:
        machines = json.load(f)
        
    machines_js = json.dumps(machines, indent=2, ensure_ascii=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bookends Hospitality — Machine QR Dashboard</title>
    
    <!-- Fonts & Icons -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    
    <!-- QRCode.js -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
    
    <!-- SheetJS (xlsx) for Excel Export -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    
    <!-- jsPDF and AutoTable for PDF Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js"></script>
    
    <!-- JSZip & FileSaver for bulk zip download -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

    <style>
        :root {{
            --bg-primary: #0b132b;
            --bg-secondary: #1c2541;
            --bg-tertiary: #3a506b;
            --accent: #48cae4;
            --accent-hover: #90e0ef;
            --text-main: #f8f9fa;
            --text-muted: #afbecf;
            --border-color: rgba(255, 255, 255, 0.08);
            --glass-bg: rgba(28, 37, 65, 0.7);
            --glass-border: rgba(255, 255, 255, 0.05);
            --shadow-premium: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            --card-brand-blue: #1f4e79;
        }}

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-primary);
            color: var(--text-main);
            min-height: 100vh;
            overflow-x: hidden;
            display: flex;
            flex-direction: column;
        }}

        /* Custom Scrollbar */
        ::-webkit-scrollbar {{
            width: 8px;
            height: 8px;
        }}
        ::-webkit-scrollbar-track {{
            background: var(--bg-primary);
        }}
        ::-webkit-scrollbar-thumb {{
            background: var(--bg-tertiary);
            border-radius: 4px;
        }}
        ::-webkit-scrollbar-thumb:hover {{
            background: var(--accent);
        }}

        /* ==================== DASHBOARD VIEW ==================== */
        header {{
            background: var(--glass-bg);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-bottom: 1px solid var(--glass-border);
            padding: 1.5rem 2rem;
            position: sticky;
            top: 0;
            z-index: 100;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: var(--shadow-premium);
        }}

        .brand {{
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }}

        .brand-logo {{
            background: linear-gradient(135deg, var(--card-brand-blue), var(--accent));
            width: 45px;
            height: 45px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 800;
            font-size: 1.25rem;
            box-shadow: 0 4px 15px rgba(31, 78, 121, 0.4);
            font-family: 'Outfit', sans-serif;
        }}

        .brand-text h1 {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.5rem;
            font-weight: 700;
            letter-spacing: 0.5px;
            background: linear-gradient(to right, #ffffff, var(--accent));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .brand-text p {{
            font-size: 0.75rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1.5px;
            margin-top: 2px;
        }}

        .header-actions {{
            display: flex;
            gap: 1rem;
        }}

        .btn {{
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.6rem 1.2rem;
            border-radius: 8px;
            font-weight: 500;
            font-size: 0.9rem;
            cursor: pointer;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
            border: none;
            outline: none;
            text-decoration: none;
        }}

        .btn-primary {{
            background: linear-gradient(135deg, var(--card-brand-blue), #2b6fa6);
            color: white;
            box-shadow: 0 4px 12px rgba(31, 78, 121, 0.3);
        }}

        .btn-primary:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 18px rgba(31, 78, 121, 0.4);
        }}

        .btn-accent {{
            background: var(--accent);
            color: var(--bg-primary);
            font-weight: 600;
        }}

        .btn-accent:hover {{
            background: var(--accent-hover);
            transform: translateY(-2px);
        }}

        .btn-outline {{
            background: transparent;
            border: 1px solid var(--border-color);
            color: var(--text-main);
        }}

        .btn-outline:hover {{
            background: rgba(255, 255, 255, 0.05);
            border-color: var(--text-muted);
        }}

        main {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem;
            width: 100%;
            flex-grow: 1;
        }}

        /* Toolbar controls */
        .controls-card {{
            background: var(--glass-bg);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            border: 1px solid var(--glass-border);
            border-radius: 16px;
            padding: 1.5rem;
            margin-bottom: 2rem;
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
            box-shadow: var(--shadow-premium);
        }}

        .search-row {{
            display: flex;
            gap: 1rem;
            flex-wrap: wrap;
        }}

        .search-wrapper {{
            flex-grow: 1;
            position: relative;
            min-width: 280px;
        }}

        .search-wrapper i {{
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            font-size: 1.1rem;
        }}

        .search-input {{
            width: 100%;
            padding: 0.75rem 1rem 0.75rem 2.8rem;
            background: rgba(0, 0, 0, 0.2);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            color: white;
            font-size: 0.95rem;
            transition: all 0.2s;
        }}

        .search-input:focus {{
            border-color: var(--accent);
            outline: none;
            box-shadow: 0 0 0 2px rgba(72, 202, 228, 0.2);
        }}

        .export-actions {{
            display: flex;
            gap: 0.75rem;
            flex-wrap: wrap;
        }}

        /* Tabs styling */
        .tabs {{
            display: flex;
            gap: 0.5rem;
            overflow-x: auto;
            padding-bottom: 5px;
            border-bottom: 1px solid var(--border-color);
        }}

        .tab-btn {{
            background: transparent;
            border: none;
            color: var(--text-muted);
            padding: 0.6rem 1.2rem;
            font-size: 0.9rem;
            font-weight: 500;
            cursor: pointer;
            white-space: nowrap;
            position: relative;
            transition: color 0.2s;
            border-radius: 6px;
        }}

        .tab-btn:hover {{
            color: white;
            background: rgba(255, 255, 255, 0.03);
        }}

        .tab-btn.active {{
            color: var(--accent);
            font-weight: 600;
            background: rgba(72, 202, 228, 0.08);
        }}

        /* Grid & Cards */
        .machines-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
            gap: 1.5rem;
        }}

        .machine-card {{
            background: var(--glass-bg);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            border: 1px solid var(--glass-border);
            border-radius: 16px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            position: relative;
        }}

        .machine-card:hover {{
            transform: translateY(-5px);
            border-color: rgba(72, 202, 228, 0.3);
            box-shadow: 0 12px 30px rgba(0,0,0,0.3);
        }}

        .card-image-section {{
            height: 180px;
            background: rgba(0, 0, 0, 0.3);
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            border-bottom: 1px solid var(--border-color);
        }}

        .card-image-section img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
        }}

        .image-placeholder {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 0.5rem;
            color: var(--text-muted);
        }}

        .image-placeholder i {{
            font-size: 2.5rem;
            opacity: 0.7;
        }}

        .badge-category {{
            position: absolute;
            top: 1rem;
            left: 1rem;
            background: rgba(11, 19, 43, 0.75);
            backdrop-filter: blur(4px);
            color: var(--accent);
            padding: 0.3rem 0.75rem;
            border-radius: 20px;
            font-size: 0.75rem;
            font-weight: 600;
            border: 1px solid rgba(72, 202, 228, 0.3);
        }}

        .card-code-banner {{
            background: var(--card-brand-blue);
            color: white;
            text-align: center;
            font-weight: 700;
            padding: 0.5rem;
            font-size: 1.1rem;
            letter-spacing: 1px;
            font-family: 'Outfit', sans-serif;
        }}

        .card-body {{
            padding: 1.25rem;
            display: flex;
            flex-direction: column;
            gap: 1rem;
            flex-grow: 1;
        }}

        .card-title-row {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
        }}

        .card-title {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.15rem;
            font-weight: 600;
            line-height: 1.3;
        }}

        .card-qty {{
            background: rgba(255,255,255,0.08);
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            font-size: 0.8rem;
            white-space: nowrap;
            color: var(--text-muted);
        }}

        .card-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.85rem;
        }}

        .card-table tr {{
            border-bottom: 1px solid rgba(255,255,255,0.04);
        }}

        .card-table tr:last-child {{
            border-bottom: none;
        }}

        .card-table td {{
            padding: 0.4rem 0;
            vertical-align: top;
        }}

        .card-table td.label {{
            color: var(--text-muted);
            width: 40%;
            font-weight: 500;
        }}

        .card-table td.value {{
            color: var(--text-main);
            font-weight: 400;
        }}

        .card-footer {{
            padding: 1rem 1.25rem;
            border-top: 1px solid var(--border-color);
            background: rgba(0, 0, 0, 0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .qr-thumbnail-wrapper {{
            width: 50px;
            height: 50px;
            background: white;
            border-radius: 6px;
            padding: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 4px 10px rgba(0,0,0,0.15);
            transition: transform 0.2s;
        }}

        .qr-thumbnail-wrapper:hover {{
            transform: scale(1.1);
        }}

        .qr-thumbnail-wrapper img {{
            max-width: 100%;
            max-height: 100%;
        }}

        .card-actions {{
            display: flex;
            gap: 0.5rem;
        }}

        .icon-btn {{
            width: 34px;
            height: 34px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            cursor: pointer;
            transition: all 0.2s;
        }}

        .icon-btn:hover {{
            background: rgba(255, 255, 255, 0.12);
            color: white;
            border-color: var(--text-muted);
        }}

        .icon-btn.edit-btn:hover {{
            color: var(--accent);
            border-color: var(--accent);
        }}

        .icon-btn.qr-btn:hover {{
            color: #2ec4b6;
            border-color: #2ec4b6;
        }}

        /* Status badge styling */
        .status-badge {{
            display: inline-block;
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .status-done {{
            background: rgba(46, 196, 182, 0.15);
            color: #2ec4b6;
        }}
        .status-pending {{
            background: rgba(255, 159, 67, 0.15);
            color: #ff9f43;
        }}
        .status-other {{
            background: rgba(255,255,255,0.08);
            color: var(--text-muted);
        }}

        /* Count Tracker Summary */
        .count-summary {{
            margin-bottom: 1.5rem;
            font-size: 0.9rem;
            color: var(--text-muted);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        /* ==================== MOBILE CARD VIEW (SAMPLE DESIGN) ==================== */
        #mobile-card-container {{
            display: none;
            width: 100%;
            max-width: 480px;
            margin: 0 auto;
            background: white;
            color: #333333;
            min-height: 100vh;
            flex-direction: column;
            box-shadow: 0 4px 30px rgba(0,0,0,0.15);
            font-family: 'Inter', sans-serif;
            position: relative;
        }}

        .mobile-header {{
            background: var(--card-brand-blue);
            color: white;
            padding: 1.25rem 1rem;
            text-align: center;
        }}

        .mobile-header h2 {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.6rem;
            font-weight: 700;
            letter-spacing: 0.5px;
            margin-bottom: 0.25rem;
        }}

        .mobile-header p {{
            font-size: 0.85rem;
            opacity: 0.9;
            font-weight: 400;
        }}

        .mobile-image-frame {{
            padding: 1rem;
            background: white;
        }}

        .mobile-image-wrapper {{
            border: 2px solid var(--card-brand-blue);
            width: 100%;
            height: 320px;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
            background: #f8f9fa;
        }}

        .mobile-image-wrapper img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
        }}

        .mobile-image-placeholder {{
            display: flex;
            flex-direction: column;
            align-items: center;
            color: #aaaaaa;
            gap: 0.5rem;
        }}

        .mobile-image-placeholder i {{
            font-size: 3rem;
        }}

        .mobile-code-banner {{
            background: var(--card-brand-blue);
            color: white;
            text-align: center;
            font-weight: 700;
            padding: 0.6rem;
            font-size: 1.3rem;
            letter-spacing: 1.5px;
            font-family: 'Outfit', sans-serif;
        }}

        .mobile-table-wrapper {{
            padding: 0.5rem 1rem;
            flex-grow: 1;
        }}

        .mobile-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 0.5rem;
        }}

        .mobile-table tr:nth-child(odd) {{
            background-color: #f8f9fa;
        }}

        .mobile-table td {{
            padding: 0.75rem 1rem;
            font-size: 0.95rem;
            vertical-align: middle;
        }}

        .mobile-table td.label {{
            font-weight: 700;
            color: #555555;
            width: 40%;
        }}

        .mobile-table td.value {{
            color: #111111;
        }}

        .mobile-footer {{
            padding: 1.25rem;
            text-align: center;
            font-size: 0.85rem;
            color: #666666;
            border-top: 1px solid #eeeeee;
            background: #fafafa;
        }}

        .back-to-dashboard-btn {{
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            margin-top: 1rem;
            padding: 0.5rem 1rem;
            border-radius: 6px;
            font-size: 0.8rem;
            background: #eeeeee;
            color: #333333;
            border: none;
            cursor: pointer;
            text-decoration: none;
            font-weight: 500;
            transition: background 0.2s;
        }}

        .back-to-dashboard-btn:hover {{
            background: #dddddd;
        }}

        /* ==================== MODAL DIALOGS ==================== */
        .modal-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.75);
            backdrop-filter: blur(4px);
            z-index: 1000;
            display: none;
            align-items: center;
            justify-content: center;
            padding: 1.5rem;
        }}

        .modal-card {{
            background: var(--bg-secondary);
            border: 1px solid var(--glass-border);
            border-radius: 16px;
            max-width: 500px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            box-shadow: var(--shadow-premium);
            display: flex;
            flex-direction: column;
        }}

        .modal-header {{
            padding: 1.25rem 1.5rem;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .modal-header h3 {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.25rem;
            font-weight: 600;
        }}

        .modal-close {{
            background: transparent;
            border: none;
            color: var(--text-muted);
            cursor: pointer;
            font-size: 1.25rem;
        }}

        .modal-close:hover {{
            color: white;
        }}

        .modal-body {{
            padding: 1.5rem;
            display: flex;
            flex-direction: column;
            gap: 1.25rem;
        }}

        .form-group {{
            display: flex;
            flex-direction: column;
            gap: 0.5rem;
        }}

        .form-group label {{
            font-size: 0.85rem;
            font-weight: 500;
            color: var(--text-muted);
        }}

        .form-input {{
            padding: 0.65rem 0.85rem;
            background: rgba(0, 0, 0, 0.2);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            color: white;
            font-size: 0.9rem;
        }}

        .form-input:focus {{
            border-color: var(--accent);
            outline: none;
        }}

        .modal-footer {{
            padding: 1.25rem 1.5rem;
            border-top: 1px solid var(--border-color);
            display: flex;
            justify-content: flex-end;
            gap: 0.75rem;
        }}

        .image-preview-box {{
            width: 100%;
            height: 150px;
            border: 2px dashed var(--border-color);
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            background: rgba(0,0,0,0.1);
            position: relative;
        }}

        .image-preview-box img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
        }}

        /* QR Show modal */
        #qr-display-container {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 1.5rem;
            padding: 1rem;
        }}

        #qr-display-image {{
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            display: flex;
            justify-content: center;
            align-items: center;
        }}

        .modal-qr-info {{
            text-align: center;
        }}

        .modal-qr-info h4 {{
            font-family: 'Outfit', sans-serif;
            font-size: 1.2rem;
            margin-bottom: 0.25rem;
        }}

        .modal-qr-info p {{
            color: var(--text-muted);
            font-size: 0.85rem;
        }}

        /* Print QR Sheet style (hidden on screen) */
        #print-qr-sheet {{
            display: none;
        }}

        @media print {{
            body * {{
                visibility: hidden;
            }}
            #mobile-card-container, #mobile-card-container * {{
                visibility: visible;
            }}
            #mobile-card-container {{
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                box-shadow: none;
            }}
            .back-to-dashboard-btn {{
                display: none;
            }}
        }}

        /* Responsiveness */
        @media (max-width: 768px) {{
            header {{
                padding: 1rem;
                flex-direction: column;
                gap: 1rem;
                align-items: stretch;
            }}
            .header-actions {{
                justify-content: space-between;
            }}
            main {{
                padding: 1rem;
            }}
            .controls-card {{
                padding: 1rem;
            }}
            .export-actions {{
                width: 100%;
            }}
            .export-actions .btn {{
                flex-grow: 1;
                justify-content: center;
            }}
        }}
    </style>
</head>
<body>

    <!-- ==================== DASHBOARD VIEW CONTROLLER ==================== -->
    <div id="dashboard-wrapper">
        <header>
            <div class="brand">
                <div class="brand-logo">BH</div>
                <div class="brand-text">
                    <h1>BOOKENDS HOSPITALITY</h1>
                    <p>Central Kitchen • Bakery, Surat</p>
                </div>
            </div>
            <div class="header-actions">
                <button class="btn btn-primary" onclick="openAddModal()">
                    <i data-lucide="plus"></i> Add Machine
                </button>
            </div>
        </header>

        <main>
            <div class="controls-card">
                <div class="search-row">
                    <div class="search-wrapper">
                        <i data-lucide="search"></i>
                        <input type="text" id="search-bar" class="search-input" placeholder="Search machines by name, brand, code, or Remark..." oninput="filterMachines()">
                    </div>
                    <div class="export-actions">
                        <button class="btn btn-outline" onclick="exportExcel()">
                            <i data-lucide="file-spreadsheet"></i> Export Excel
                        </button>
                        <button class="btn btn-outline" onclick="exportPDF()">
                            <i data-lucide="file-text"></i> Export PDF
                        </button>
                        <button class="btn btn-outline" onclick="downloadAllQRs()">
                            <i data-lucide="download"></i> Download QRs (.zip)
                        </button>
                    </div>
                </div>

                <div class="tabs" id="category-tabs">
                    <button class="tab-btn active" onclick="switchCategory('All')">All Machines</button>
                    <button class="tab-btn" onclick="switchCategory('Bakery Machine List')">Bakery</button>
                    <button class="tab-btn" onclick="switchCategory('Prep Kitchen')">Prep Kitchen</button>
                    <button class="tab-btn" onclick="switchCategory('New Machines (2026)')">New (2026)</button>
                    <button class="tab-btn" onclick="switchCategory('R&D Kitchen')">R&D Kitchen</button>
                </div>
            </div>

            <div class="count-summary">
                <span id="filtered-count">Showing 0 of 0 machines</span>
            </div>

            <div class="machines-grid" id="machines-container">
                <!-- Dynamically rendered cards -->
            </div>
        </main>
    </div>

    <!-- ==================== MOBILE CARD VIEW CONTAINER ==================== -->
    <div id="mobile-card-container">
        <div class="mobile-header">
            <h2 id="m-header-title">BOOKENDS HOSPITALITY</h2>
            <p id="m-header-subtitle">Central Kitchen • Bakery, Surat</p>
        </div>
        <div class="mobile-image-frame">
            <div class="mobile-image-wrapper">
                <img id="m-machine-image" src="" alt="Machine Image" style="display:none;">
                <div id="m-image-placeholder" class="mobile-image-placeholder">
                    <i data-lucide="chef-hat"></i>
                    <span>No Image Available</span>
                </div>
            </div>
        </div>
        <div class="mobile-code-banner" id="m-machine-code">SUR-BKY-001</div>
        <div class="mobile-table-wrapper">
            <table class="mobile-table" id="m-table">
                <!-- Fields populated dynamically -->
            </table>
        </div>
        <div class="mobile-footer">
            <div>Scan → is machine ki poori detail + photo</div>
            <a href="?" class="back-to-dashboard-btn" id="m-back-btn">
                <i data-lucide="layout-dashboard"></i> Open Dashboard
            </a>
        </div>
    </div>

    <!-- ==================== ADD / EDIT MACHINE MODAL ==================== -->
    <div class="modal-overlay" id="machine-modal">
        <div class="modal-card">
            <div class="modal-header">
                <h3 id="modal-title">Add New Machine</h3>
                <button class="modal-close" onclick="closeModal('machine-modal')">&times;</button>
            </div>
            <div class="modal-body">
                <input type="hidden" id="edit-machine-idx">
                
                <div class="form-group">
                    <label for="form-category">Category</label>
                    <select id="form-category" class="form-input" onchange="suggestNextCode()">
                        <option value="Bakery Machine List">Bakery Machine List</option>
                        <option value="Prep Kitchen">Prep Kitchen</option>
                        <option value="New Machines (2026)">New Machines (2026)</option>
                        <option value="R&D Kitchen">R&D Kitchen</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label for="form-code">Equipment Code</label>
                    <input type="text" id="form-code" class="form-input" placeholder="e.g. SUR-BKY-001">
                </div>

                <div class="form-group">
                    <label for="form-name">Machine / Item Name</label>
                    <input type="text" id="form-name" class="form-input" placeholder="Enter machine name">
                </div>

                <div class="form-group">
                    <label for="form-qty">Quantity</label>
                    <input type="text" id="form-qty" class="form-input" value="1">
                </div>

                <div class="form-group">
                    <label for="form-brand">Brand</label>
                    <input type="text" id="form-brand" class="form-input" placeholder="Brand name">
                </div>

                <div class="form-group" id="form-capacity-group">
                    <label for="form-capacity">Capacity / Details</label>
                    <input type="text" id="form-capacity" class="form-input" placeholder="Capacity or detailed specs">
                </div>

                <div class="form-group" id="form-date-group">
                    <label for="form-date">Purchase Date / Remark / Use</label>
                    <input type="text" id="form-date" class="form-input" placeholder="e.g. Date or Application or Remark">
                </div>

                <div class="form-group">
                    <label>Machine Photo</label>
                    <div class="image-preview-box" id="modal-image-preview-box" style="margin-bottom: 0.5rem;">
                        <div class="image-placeholder" id="modal-image-placeholder">
                            <i data-lucide="image"></i>
                            <span>Click to upload image</span>
                        </div>
                        <img id="modal-image-preview" src="" style="display:none;">
                    </div>
                    <input type="file" id="form-image-file" accept="image/*" style="display:none;" onchange="handleImageUpload(event)">
                </div>

                <div class="form-group">
                    <label for="form-image-url">Or paste Machine Photo URL (e.g. GitHub link)</label>
                    <input type="text" id="form-image-url" class="form-input" placeholder="Paste direct image URL (https://...)" oninput="handleImageUrlInput()">
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-outline" onclick="closeModal('machine-modal')">Cancel</button>
                <button class="btn btn-accent" onclick="saveMachine()">Save Machine</button>
            </div>
        </div>
    </div>

    <!-- ==================== QR VIEW MODAL ==================== -->
    <div class="modal-overlay" id="qr-modal">
        <div class="modal-card" style="max-width: 380px;">
            <div class="modal-header">
                <h3>Machine QR Code</h3>
                <button class="modal-close" onclick="closeModal('qr-modal')">&times;</button>
            </div>
            <div class="modal-body" id="qr-display-container">
                <div id="qr-display-image"></div>
                <div class="modal-qr-info">
                    <h4 id="qr-modal-name">Machine Name</h4>
                    <p id="qr-modal-code">SUR-BKY-001</p>
                </div>
                <button class="btn btn-primary" style="width: 100%; justify-content: center;" id="download-single-qr-btn">
                    <i data-lucide="download"></i> Download QR (PNG)
                </button>
            </div>
        </div>
    </div>

    <!-- Helper hidden container for generating offscreen QRs for zip/pdf -->
    <div id="hidden-qr-temp" style="display:none; width: 256px; height: 256px; background: white; padding: 10px;"></div>

    <script>
        // Core data
        const initialMachines = {machines_js};

        // Live state loaded from localStorage or fallback to initialMachines
        let machines = [];
        let currentCategory = 'All';

        // Load custom database
        function initDB() {{
            const stored = localStorage.getItem('bookends_machines');
            if (stored) {{
                try {{
                    machines = JSON.parse(stored);
                }} catch(e) {{
                    machines = [...initialMachines];
                }}
            }} else {{
                machines = [...initialMachines];
                saveToLocalStorage();
            }}
        }}

        function saveToLocalStorage() {{
            localStorage.setItem('bookends_machines', JSON.stringify(machines));
        }}

        // Check URL for Mobile Card mode
        window.addEventListener('DOMContentLoaded', () => {{
            initDB();
            lucide.createIcons();

            const urlParams = new URLSearchParams(window.location.search);
            const machineCode = urlParams.get('machine');

            if (machineCode) {{
                showMobileCard(machineCode);
            }} else {{
                showDashboard();
            }}

            // Setup image upload trigger click
            document.getElementById('modal-image-preview-box').addEventListener('click', () => {{
                document.getElementById('form-image-file').click();
            }});
        }});

        function showDashboard() {{
            document.getElementById('dashboard-wrapper').style.display = 'block';
            document.getElementById('mobile-card-container').style.display = 'none';
            renderGrid();
        }}

        // Format data values for safe viewing
        function val(valString) {{
            return (!valString || valString === '-' || valString === 'None' || valString === 'N/A') ? '-' : valString;
        }}

        function showMobileCard(code) {{
            document.getElementById('dashboard-wrapper').style.display = 'none';
            const container = document.getElementById('mobile-card-container');
            container.style.display = 'flex';

            const m = machines.find(item => item.id.toUpperCase() === code.toUpperCase());
            if (!m) {{
                document.getElementById('m-machine-code').textContent = 'NOT FOUND';
                document.getElementById('m-table').innerHTML = `
                    <tr><td colspan="2" style="text-align:center; padding: 2rem; color: #888;">Machine with code "${{code}}" could not be found.</td></tr>
                `;
                return;
            }}

            document.getElementById('m-machine-code').textContent = m.id;

            // Load Image
            const imgEl = document.getElementById('m-machine-image');
            const placeholderEl = document.getElementById('m-image-placeholder');
            
            if (m.image) {{
                imgEl.src = m.image;
                imgEl.style.display = 'block';
                placeholderEl.style.display = 'none';
            }} else {{
                imgEl.style.display = 'none';
                placeholderEl.style.display = 'flex';
            }}

            // Populate table based on category exactly matches SAMPLE_card_BOOKENDS_3.png
            let rowsHtml = '';
            
            if (m.category === 'Bakery Machine List') {{
                rowsHtml = `
                    <tr>
                        <td class="label">Machine Name</td>
                        <td class="value">${{m.name}}</td>
                    </tr>
                    <tr>
                        <td class="label">Brand</td>
                        <td class="value">${{val(m.brand)}}</td>
                    </tr>
                    <tr>
                        <td class="label">Capacity</td>
                        <td class="value">${{val(m.capacity)}}</td>
                    </tr>
                    <tr>
                        <td class="label">Purchase Date</td>
                        <td class="value">${{val(m.purchase_date)}}</td>
                    </tr>
                    <tr>
                        <td class="label">Quantity</td>
                        <td class="value">${{m.qty || 1}}</td>
                    </tr>
                `;
            }} else if (m.category === 'Prep Kitchen' || m.category === 'R&D Kitchen') {{
                rowsHtml = `
                    <tr>
                        <td class="label">Machine Name</td>
                        <td class="value">${{m.name}}</td>
                    </tr>
                    <tr>
                        <td class="label">Brand</td>
                        <td class="value">${{val(m.brand)}}</td>
                    </tr>
                    <tr>
                        <td class="label">Use/Application</td>
                        <td class="value">${{val(m.use)}}</td>
                    </tr>
                    <tr>
                        <td class="label">Quantity</td>
                        <td class="value">${{m.qty || 1}}</td>
                    </tr>
                    <tr>
                        <td class="label">Contact No.</td>
                        <td class="value">${{val(m.contact)}}</td>
                    </tr>
                `;
            }} else {{
                // New Machines or other
                rowsHtml = `
                    <tr>
                        <td class="label">Machine Name</td>
                        <td class="value">${{m.name}}</td>
                    </tr>
                    <tr>
                        <td class="label">Quantity</td>
                        <td class="value">${{m.qty || 1}}</td>
                    </tr>
                    <tr>
                        <td class="label">Status/Remark</td>
                        <td class="value">${{val(m.remark)}}</td>
                    </tr>
                `;
            }}

            document.getElementById('m-table').innerHTML = rowsHtml;
            lucide.createIcons();
        }}

        // Render dashboard grid of cards
        function renderGrid() {{
            const searchVal = document.getElementById('search-bar').value.toLowerCase().trim();
            const container = document.getElementById('machines-container');
            container.innerHTML = '';

            let filtered = machines;

            // Filter Category
            if (currentCategory !== 'All') {{
                filtered = filtered.filter(item => item.category === currentCategory);
            }}

            // Filter Search Text
            if (searchVal !== '') {{
                filtered = filtered.filter(item => {{
                    return item.name.toLowerCase().includes(searchVal) || 
                           item.id.toLowerCase().includes(searchVal) || 
                           (item.brand && item.brand.toLowerCase().includes(searchVal)) ||
                           (item.use && item.use.toLowerCase().includes(searchVal)) ||
                           (item.remark && item.remark.toLowerCase().includes(searchVal));
                }});
            }}

            // Update Counter
            document.getElementById('filtered-count').textContent = `Showing ${{filtered.length}} of ${{machines.length}} machines`;

            if (filtered.length === 0) {{
                container.innerHTML = `
                    <div style="grid-column: 1/-1; text-align: center; padding: 4rem 2rem; color: var(--text-muted);">
                        <i data-lucide="search-code" style="font-size: 3rem; margin-bottom: 1rem; opacity: 0.5;"></i>
                        <p>No machines match your search or filter criteria.</p>
                    </div>
                `;
                lucide.createIcons();
                return;
            }}

            filtered.forEach((m, idx) => {{
                // Render card
                const card = document.createElement('div');
                card.className = 'machine-card';
                card.dataset.id = m.id;

                let tableRows = '';
                if (m.category === 'Bakery Machine List') {{
                    tableRows = `
                        <tr><td class="label">Brand:</td><td class="value">${{val(m.brand)}}</td></tr>
                        <tr><td class="label">Capacity:</td><td class="value">${{val(m.capacity)}}</td></tr>
                        <tr><td class="label">Purchased:</td><td class="value">${{val(m.purchase_date)}}</td></tr>
                    `;
                }} else if (m.category === 'Prep Kitchen' || m.category === 'R&D Kitchen') {{
                    tableRows = `
                        <tr><td class="label">Brand:</td><td class="value">${{val(m.brand)}}</td></tr>
                        <tr><td class="label">Application:</td><td class="value">${{val(m.use)}}</td></tr>
                        <tr><td class="label">Contact:</td><td class="value">${{val(m.contact)}}</td></tr>
                    `;
                }} else {{
                    let statusClass = 'status-other';
                    if (m.remark && m.remark.toLowerCase().includes('done')) statusClass = 'status-done';
                    else if (m.remark && m.remark.toLowerCase().includes('pending')) statusClass = 'status-pending';
                    
                    tableRows = `
                        <tr><td class="label">Remark:</td><td class="value"><span class="status-badge ${{statusClass}}">${{val(m.remark)}}</span></td></tr>
                    `;
                }}

                // Generate a temporary scan URL
                const scanUrl = `${{window.location.origin}}${{window.location.pathname}}?machine=${{m.id}}`;

                card.innerHTML = `
                    <div class="card-image-section">
                        <span class="badge-category">${{m.category}}</span>
                        ${{m.image ? `<img src="${{m.image}}" alt="${{m.name}}">` : `
                        <div class="image-placeholder">
                            <i data-lucide="chef-hat"></i>
                            <span>No Photo</span>
                        </div>
                        `}}
                    </div>
                    <div class="card-code-banner">${{m.id}}</div>
                    <div class="card-body">
                        <div class="card-title-row">
                            <div class="card-title">${{m.name}}</div>
                            <div class="card-qty">Qty: ${{m.qty || 1}}</div>
                        </div>
                        <table class="card-table">
                            ${{tableRows}}
                        </table>
                    </div>
                    <div class="card-footer">
                        <div class="qr-thumbnail-wrapper" onclick="showQR('${{m.id}}', '${{m.name.replace(/'/g, "\\'")}}')" title="Click to view & download QR">
                            <div class="qr-thumbnail" id="qr-thumb-${{m.id}}"></div>
                        </div>
                        <div class="card-actions">
                            <button class="icon-btn edit-btn" onclick="openEditModal('${{m.id}}')" title="Edit details & photo">
                                <i data-lucide="edit"></i>
                            </button>
                            <button class="icon-btn qr-btn" onclick="showQR('${{m.id}}', '${{m.name.replace(/'/g, "\\'")}}')" title="View QR Code">
                                <i data-lucide="qr-code"></i>
                            </button>
                            <button class="icon-btn" onclick="deleteMachine('${{m.id}}')" title="Delete entry" style="color: #ff6b6b; border-color: rgba(255,107,107,0.2);">
                                <i data-lucide="trash-2"></i>
                            </button>
                        </div>
                    </div>
                `;

                container.appendChild(card);

                // Auto render the small thumbnail QR code
                new QRCode(document.getElementById(`qr-thumb-${{m.id}}`), {{
                    text: scanUrl,
                    width: 42,
                    height: 42,
                    correctLevel: QRCode.CorrectLevel.M
                }});
            }});

            lucide.createIcons();
        }}

        // Switch filter category tabs
        function switchCategory(cat) {{
            currentCategory = cat;
            
            // Toggle active classes
            const btns = document.querySelectorAll('#category-tabs .tab-btn');
            btns.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');

            renderGrid();
        }}

        function filterMachines() {{
            renderGrid();
        }}

        // Modal triggers
        function openAddModal() {{
            document.getElementById('modal-title').textContent = 'Add New Machine';
            document.getElementById('edit-machine-idx').value = '';
            
            // clear form inputs
            document.getElementById('form-name').value = '';
            document.getElementById('form-qty').value = '1';
            document.getElementById('form-brand').value = '';
            document.getElementById('form-capacity').value = '';
            document.getElementById('form-date').value = '';
            document.getElementById('modal-image-preview').src = '';
            document.getElementById('modal-image-preview').style.display = 'none';
            document.getElementById('modal-image-placeholder').style.display = 'flex';
            document.getElementById('form-image-file').value = '';
            document.getElementById('form-image-url').value = '';
            
            // suggest code
            suggestNextCode();

            document.getElementById('machine-modal').style.display = 'flex';
        }}

        function suggestNextCode() {{
            const cat = document.getElementById('form-category').value;
            let prefix = 'SUR-BKY';
            if (cat === 'Prep Kitchen') prefix = 'SUR-PRP';
            else if (cat === 'New Machines (2026)') prefix = 'SUR-NEW';
            else if (cat === 'R&D Kitchen') prefix = 'SUR-RD';

            const matchItems = machines.filter(item => item.id.startsWith(prefix));
            let maxNum = 0;
            matchItems.forEach(item => {{
                const numStr = item.id.replace(prefix + '-', '');
                const num = parseInt(numStr);
                if (!isNaN(num) && num > maxNum) maxNum = num;
            }});

            const nextCode = `${{prefix}}-${{String(maxNum + 1).padStart(3, '0')}}`;
            document.getElementById('form-code').value = nextCode;

            // adjust labels based on category
            const capGroup = document.getElementById('form-capacity-group');
            const dateLabel = document.querySelector('#form-date-group label');
            const dateInput = document.getElementById('form-date');

            if (cat === 'Bakery Machine List') {{
                capGroup.style.display = 'flex';
                document.querySelector('#form-capacity-group label').textContent = 'Capacity';
                dateLabel.textContent = 'Purchase Date';
                dateInput.placeholder = 'e.g. 12-Nov-2025';
            }} else if (cat === 'Prep Kitchen' || cat === 'R&D Kitchen') {{
                capGroup.style.display = 'flex';
                document.querySelector('#form-capacity-group label').textContent = 'Use / Application';
                dateLabel.textContent = 'Contact Number';
                dateInput.placeholder = 'Supplier or Service contact';
            }} else {{
                capGroup.style.display = 'none';
                dateLabel.textContent = 'Remark / Status';
                dateInput.placeholder = 'e.g. Pending / Done';
            }}
        }}

        function openEditModal(code) {{
            const m = machines.find(item => item.id === code);
            if (!m) return;

            document.getElementById('modal-title').textContent = 'Edit Machine Details';
            document.getElementById('edit-machine-idx').value = code;

            document.getElementById('form-category').value = m.category;
            document.getElementById('form-code').value = m.id;
            document.getElementById('form-name').value = m.name;
            document.getElementById('form-qty').value = m.qty || '1';
            document.getElementById('form-brand').value = m.brand || '';

            // adjust fields & labels
            const cat = m.category;
            const capGroup = document.getElementById('form-capacity-group');
            const dateLabel = document.querySelector('#form-date-group label');
            const dateInput = document.getElementById('form-date');

            if (cat === 'Bakery Machine List') {{
                capGroup.style.display = 'flex';
                document.querySelector('#form-capacity-group label').textContent = 'Capacity';
                document.getElementById('form-capacity').value = m.capacity || '';
                dateLabel.textContent = 'Purchase Date';
                dateInput.value = m.purchase_date || '';
            }} else if (cat === 'Prep Kitchen' || cat === 'R&D Kitchen') {{
                capGroup.style.display = 'flex';
                document.querySelector('#form-capacity-group label').textContent = 'Use / Application';
                document.getElementById('form-capacity').value = m.use || '';
                dateLabel.textContent = 'Contact Number';
                dateInput.value = m.contact || '';
            }} else {{
                capGroup.style.display = 'none';
                dateLabel.textContent = 'Remark / Status';
                dateInput.value = m.remark || '';
            }}

            const preview = document.getElementById('modal-image-preview');
            const placeholder = document.getElementById('modal-image-placeholder');

            if (m.image) {{
                preview.src = m.image;
                preview.style.display = 'block';
                placeholder.style.display = 'none';
                if (!m.image.startsWith('data:image')) {{
                    document.getElementById('form-image-url').value = m.image;
                }} else {{
                    document.getElementById('form-image-url').value = '';
                }}
            }} else {{
                preview.src = '';
                preview.style.display = 'none';
                placeholder.style.display = 'flex';
                document.getElementById('form-image-url').value = '';
            }}

            document.getElementById('machine-modal').style.display = 'flex';
        }}

        function closeModal(id) {{
            document.getElementById(id).style.display = 'none';
        }}

        // Handle Image Upload with Base64 encoding
        let base64Image = '';
        function handleImageUpload(e) {{
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(event) {{
                base64Image = event.target.result;
                const preview = document.getElementById('modal-image-preview');
                const placeholder = document.getElementById('modal-image-placeholder');
                
                preview.src = base64Image;
                preview.style.display = 'block';
                placeholder.style.display = 'none';
                // Clear URL input if file is uploaded
                document.getElementById('form-image-url').value = '';
            }};
            reader.readAsDataURL(file);
        }}

        function handleImageUrlInput() {{
            const url = document.getElementById('form-image-url').value.trim();
            const preview = document.getElementById('modal-image-preview');
            const placeholder = document.getElementById('modal-image-placeholder');

            if (url) {{
                preview.src = url;
                preview.style.display = 'block';
                placeholder.style.display = 'none';
                // Clear file input if URL is pasted
                document.getElementById('form-image-file').value = '';
            }} else {{
                preview.src = '';
                preview.style.display = 'none';
                placeholder.style.display = 'flex';
            }}
        }}

        function saveMachine() {{
            const code = document.getElementById('form-code').value.trim();
            const cat = document.getElementById('form-category').value;
            const name = document.getElementById('form-name').value.trim();
            const qty = document.getElementById('form-qty').value.trim();
            const brand = document.getElementById('form-brand').value.trim();
            const capVal = document.getElementById('form-capacity').value.trim();
            const dateVal = document.getElementById('form-date').value.trim();

            if (!code || !name) {{
                alert('Please enter Equipment Code and Machine Name.');
                return;
            }}

            const editCode = document.getElementById('edit-machine-idx').value;
            
            // Build the machine entry
            const machineObj = {{
                id: code,
                category: cat,
                name: name,
                qty: qty || '1',
                brand: brand || '-',
                details: ''
            }};

            if (cat === 'Bakery Machine List') {{
                machineObj.capacity = capVal || '-';
                machineObj.purchase_date = dateVal || '-';
                machineObj.details = `Equipment Code: ${{code}}. Brand: ${{brand || '-'}}. Capacity: ${{capVal || '-'}}. Purchase Date: ${{dateVal || '-'}}`;
            }} else if (cat === 'Prep Kitchen' || cat === 'R&D Kitchen') {{
                machineObj.use = capVal || '-';
                machineObj.contact = dateVal || '-';
                machineObj.details = `Use: ${{capVal || '-'}}. Brand: ${{brand || '-'}}. Contact: ${{dateVal || '-'}}`;
            }} else {{
                machineObj.remark = dateVal || '-';
                machineObj.details = `Remark: ${{dateVal || '-'}}`;
            }}

            // Retain/apply base64 image or URL
            const urlInputVal = document.getElementById('form-image-url').value.trim();
            const previewSrc = document.getElementById('modal-image-preview').src;

            if (urlInputVal) {{
                machineObj.image = urlInputVal;
            }} else if (previewSrc && previewSrc.startsWith('data:image')) {{
                machineObj.image = previewSrc;
            }} else if (editCode) {{
                // reuse existing image if editing
                const existing = machines.find(x => x.id === editCode);
                if (existing && existing.image) machineObj.image = existing.image;
            }}

            if (editCode) {{
                // Update
                const idx = machines.findIndex(x => x.id === editCode);
                if (idx !== -1) {{
                    machines[idx] = machineObj;
                }}
            }} else {{
                // Add new (verify duplicate code first)
                if (machines.some(x => x.id.toUpperCase() === code.toUpperCase())) {{
                    alert(`A machine with code "${{code}}" already exists.`);
                    return;
                }}
                machines.unshift(machineObj);
            }}

            saveToLocalStorage();
            closeModal('machine-modal');
            renderGrid();
        }}

        function deleteMachine(code) {{
            if (confirm(`Are you sure you want to delete machine "${{code}}"?`)) {{
                machines = machines.filter(x => x.id !== code);
                saveToLocalStorage();
                renderGrid();
            }}
        }}

        // ==================== SINGLE QR MODAL SHOW & DOWNLOAD ====================
        function showQR(code, name) {{
            const container = document.getElementById('qr-display-image');
            container.innerHTML = '';

            const scanUrl = `${{window.location.origin}}${{window.location.pathname}}?machine=${{code}}`;

            new QRCode(container, {{
                text: scanUrl,
                width: 200,
                height: 200,
                correctLevel: QRCode.CorrectLevel.H
            }});

            document.getElementById('qr-modal-name').textContent = name;
            document.getElementById('qr-modal-code').textContent = code;

            const dlBtn = document.getElementById('download-single-qr-btn');
            dlBtn.onclick = () => {{
                setTimeout(() => {{
                    const img = container.querySelector('img');
                    if (img) {{
                        const link = document.createElement('a');
                        link.download = `${{code}}_${{name.replace(/\s+/g, '_')}}.png`;
                        link.href = img.src;
                        link.click();
                    }}
                }}, 200);
            }};

            document.getElementById('qr-modal').style.display = 'flex';
        }}

        // ==================== EXPORT DATA TO EXCEL ====================
        function exportExcel() {{
            let exportData = [];
            
            // Format columns nicely
            machines.forEach(m => {{
                if (m.category === 'Bakery Machine List') {{
                    exportData.push({{
                        'Category': m.category,
                        'Equipment Code': m.id,
                        'Machine Name': m.name,
                        'Qty': m.qty,
                        'Brand': m.brand,
                        'Capacity': m.capacity,
                        'Purchase Date': m.purchase_date,
                        'Scan URL': `${{window.location.origin}}${{window.location.pathname}}?machine=${{m.id}}`
                    }});
                }} else if (m.category === 'Prep Kitchen' || m.category === 'R&D Kitchen') {{
                    exportData.push({{
                        'Category': m.category,
                        'Equipment Code': m.id,
                        'Machine Name': m.name,
                        'Qty': m.qty,
                        'Brand': m.brand,
                        'Use / Application': m.use,
                        'Contact / Supplier': m.contact,
                        'Scan URL': `${{window.location.origin}}${{window.location.pathname}}?machine=${{m.id}}`
                    }});
                }} else {{
                    exportData.push({{
                        'Category': m.category,
                        'Equipment Code': m.id,
                        'Machine Name': m.name,
                        'Qty': m.qty,
                        'Remark / Status': m.remark,
                        'Scan URL': `${{window.location.origin}}${{window.location.pathname}}?machine=${{m.id}}`
                    }});
                }}
            }});

            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Machines Master Data");
            XLSX.writeFile(workbook, "Bookends_Hospitality_Machines.xlsx");
        }}

        // ==================== EXPORT PDF WITH PRINT QR ====================
        async function exportPDF() {{
            const {{ jsPDF }} = window.jspdf;
            const doc = new jsPDF('p', 'mm', 'a4');
            
            doc.setFillColor(31, 78, 121);
            doc.rect(0, 0, 210, 30, 'F');
            doc.setTextColor(255, 255, 255);
            doc.setFont('Helvetica', 'bold');
            doc.setFontSize(20);
            doc.text("BOOKENDS HOSPITALITY", 15, 18);
            doc.setFontSize(10);
            doc.setFont('Helvetica', 'normal');
            doc.text("Central Kitchen • Bakery, Surat — QR Print Catalog", 15, 24);

            let tableData = [];
            
            // We'll gather the data rows
            machines.forEach((m, idx) => {{
                tableData.push([
                    m.id,
                    m.name,
                    m.qty || 1,
                    m.category,
                    m.brand || m.remark || '-'
                ]);
            }});

            doc.autoTable({{
                head: [['Code', 'Machine Name', 'Qty', 'Category', 'Brand / Remark']],
                body: tableData,
                startY: 40,
                theme: 'striped',
                headStyles: {{ fillColor: [31, 78, 121] }},
                styles: {{ fontSize: 9, cellPadding: 4 }},
                margin: {{ top: 40 }},
                didDrawPage: function (data) {{
                    // Footer
                    doc.setFontSize(8);
                    doc.setTextColor(120);
                    doc.text("Page " + doc.internal.getNumberOfPages() + " | Bookends Hospitality Surat", 15, 285);
                }}
            }});

            // Let's also print standard scannable labels on a separate sheet!
            doc.addPage();
            doc.setFillColor(31, 78, 121);
            doc.rect(0, 0, 210, 15, 'F');
            doc.setTextColor(255, 255, 255);
            doc.setFontSize(14);
            doc.setFont('Helvetica', 'bold');
            doc.text("SCANABLE LABELS (QR CODES FOR PRINTING)", 15, 10);

            // Print 3x6 grid of QR labels on A4 page
            let x = 15;
            let y = 25;
            let qrSize = 35;
            let stepX = 65;
            let stepY = 45;

            // Loop and draw QRs
            const tempDiv = document.getElementById('hidden-qr-temp');

            for (let i = 0; i < Math.min(24, machines.length); i++) {{
                const m = machines[i];
                tempDiv.innerHTML = '';
                const scanUrl = `${{window.location.origin}}${{window.location.pathname}}?machine=${{m.id}}`;
                
                // create QR code inside temp div synchronously
                new QRCode(tempDiv, {{
                    text: scanUrl,
                    width: 120,
                    height: 120,
                    correctLevel: QRCode.CorrectLevel.M
                }});

                // Wait tiny bit for the QR library canvas to render
                await new Promise(r => setTimeout(r, 10));

                const qrImg = tempDiv.querySelector('img');
                if (qrImg && qrImg.src) {{
                    // draw box border
                    doc.setDrawColor(200, 200, 200);
                    doc.rect(x - 2, y - 2, qrSize + 15, qrSize + 8);

                    // Add QR Code image
                    doc.addImage(qrImg.src, 'PNG', x, y, qrSize, qrSize);
                    
                    // Add details text
                    doc.setTextColor(0, 0, 0);
                    doc.setFontSize(7);
                    doc.setFont('Helvetica', 'bold');
                    doc.text(m.id, x + qrSize + 1, y + 8);
                    
                    doc.setFont('Helvetica', 'normal');
                    // Wrap text for machine name
                    const splitName = doc.splitTextToSize(m.name, 22);
                    doc.text(splitName, x + qrSize + 1, y + 15);
                    
                    // Move coordinates
                    x += stepX;
                    if (x > 180) {{
                        x = 15;
                        y += stepY;
                    }}
                    if (y > 260 && i < machines.length - 1) {{
                        doc.addPage();
                        x = 15;
                        y = 20;
                    }}
                }}
            }}

            doc.save("Bookends_Hospitality_Print_Labels.pdf");
        }}

        // ==================== DOWNLOAD ALL QR CODES AS ZIP ====================
        async function downloadAllQRs() {{
            const zip = new JSZip();
            const tempDiv = document.getElementById('hidden-qr-temp');
            
            let btn = event.target;
            const originalText = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = `<i class="spinner"></i> Generating Zips...`;

            for (let i = 0; i < machines.length; i++) {{
                const m = machines[i];
                tempDiv.innerHTML = '';
                const scanUrl = `${{window.location.origin}}${{window.location.pathname}}?machine=${{m.id}}`;

                new QRCode(tempDiv, {{
                    text: scanUrl,
                    width: 256,
                    height: 256,
                    correctLevel: QRCode.CorrectLevel.H
                }});

                // Give it a tiny moment to draw the image
                await new Promise(r => setTimeout(r, 20));

                const qrImg = tempDiv.querySelector('img');
                if (qrImg && qrImg.src) {{
                    const base64Data = qrImg.src.split(',')[1];
                    const filename = `${{m.id}}_${{m.name.replace(/[^a-zA-Z0-9]/g, '_')}}.png`;
                    zip.file(filename, base64Data, {{base64: true}});
                }}
            }}

            zip.generateAsync({{type:"blob"}}).then(function(content) {{
                saveAs(content, "Bookends_Hospitality_QR_Codes.zip");
                btn.disabled = false;
                btn.innerHTML = originalText;
            }});
        }}
    </script>
</body>
</html>
"""
    
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("Build complete: index.html generated successfully!")

if __name__ == '__main__':
    build()
