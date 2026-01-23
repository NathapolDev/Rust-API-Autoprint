use actix_web::{get, post, web, App, HttpResponse, HttpServer, Responder};
use anyhow::{bail, Context, Result};
use lopdf::content::{Content, Operation};
use lopdf::{Document, Object};
use printers::{self, common::base::job::PrinterJobOptions};
use serde::{Deserialize, Serialize};

use std::ffi::OsString;
use std::path::Path;
use std::sync::mpsc;
use std::time::Duration;
use utoipa::{OpenApi, ToSchema};
use utoipa_swagger_ui::SwaggerUi;
use windows_service::{
    define_windows_service,
    service::{
        ServiceControl, ServiceControlAccept, ServiceExitCode, ServiceState, ServiceStatus,
        ServiceType,
    },
    service_control_handler::{self, ServiceControlHandlerResult},
    service_dispatcher,
};

const SERVICE_NAME: &str = "rust-print-api";
const SERVICE_TYPE: ServiceType = ServiceType::OWN_PROCESS;

// --- ค่าคงที่สำหรับขนาดกระดาษในหน่วย PostScript Points (1 point = 1/72 inch) ---

// A6: 105mm x 148mm
const A6_WIDTH_PTS: f32 = 297.64;
const A6_HEIGHT_PTS: f32 = 419.53;

// A4: 210mm x 297mm (ใช้เป็นขนาดอ้างอิงของเอกสารต้นฉบับ)
const A4_WIDTH_PTS: f32 = 595.28;
const A4_HEIGHT_PTS: f32 = 841.89;

/// โครงสร้างสำหรับรับข้อมูลจาก HTTP Request (JSON)
#[derive(Deserialize, ToSchema)]
#[schema(example = json!({"filename": "invoice_original.pdf", "printer_name": "Office_LaserJet"}))]
struct PrintRequest {
    /// ชื่อไฟล์ PDF ต้นฉบับที่จะค้นหาในโฟลเดอร์ ./printable_files
    filename: String,
    /// ชื่อเครื่องพิมพ์ปลายทางที่ติดตั้งในระบบ
    printer_name: String,
}

/// โครงสร้างสำหรับ Response ที่ส่งกลับไปให้ Client
#[derive(Serialize, ToSchema)]
struct ResponseMessage {
    status: String,
    message: String,
}

/// โครงสร้างสำหรับ Response รายการเครื่องพิมพ์
#[derive(Serialize, ToSchema)]
struct PrinterListResponse {
    status: String,
    printers: Vec<String>,
}

// ----------------------------------------------------------------------
//                        PDF RESIZING LOGIC (WITH SCALING)
// ----------------------------------------------------------------------

/// แปลงขนาด PDF จากไฟล์ต้นฉบับเป็น A6 และปรับมาตราส่วนเนื้อหา
fn resize_pdf_to_a6(input_path: &Path, output_path: &Path) -> Result<()> {
    let mut doc = Document::load(input_path)
        .context(format!("Failed to load PDF file: {}", input_path.display()))?;

    // คำนวณ Scale Factor (สมมติ A4 เป็นขนาดตั้งต้น)
    let scale_x = A6_WIDTH_PTS / A4_WIDTH_PTS;
    let scale_y = A6_HEIGHT_PTS / A4_HEIGHT_PTS;
    let scale_factor = scale_x.min(scale_y);

    if scale_factor > 1.0 {
        bail!("Scaling up is not handled, only scaling down to A6.");
    }

    for (_, page_id) in doc.get_pages() {
        // Modify MediaBox
        if let Ok(page) = doc.get_dictionary_mut(page_id) {
            let new_media_box = vec![
                Object::Real(0.0),
                Object::Real(0.0),
                Object::Real(A6_WIDTH_PTS),
                Object::Real(A6_HEIGHT_PTS),
            ];
            page.set("MediaBox", Object::Array(new_media_box));
        }

        // Modify content stream
        let content_data = doc.get_page_content(page_id)?;
        let mut content = Content::decode(&content_data)?;

        let matrix_op = Operation::new(
            "cm",
            vec![
                Object::Real(scale_factor),
                Object::Real(0.0),
                Object::Real(0.0),
                Object::Real(scale_factor),
                Object::Real(0.0),
                Object::Real(0.0),
            ],
        );
        content.operations.insert(0, matrix_op);

        let new_content = content.encode()?;
        doc.change_page_content(page_id, new_content)?;
    }

    doc.save(output_path)
        .context(format!("Failed to save new A6 PDF file: {}", output_path.display()))?;

    Ok(())
}

// ----------------------------------------------------------------------
//                           API HANDLER (UPDATED)
// ----------------------------------------------------------------------

/// กำหนดโครงสร้างเอกสาร OpenAPI
#[derive(OpenApi)]
#[openapi(
    paths(print_file_handler, get_printers, index),
    components(schemas(PrintRequest, ResponseMessage, PrinterListResponse)),
    tags((name = "Printing", description = "Endpoints สำหรับการดำเนินการสั่งพิมพ์ไฟล์และแปลงขนาด"))
)]
struct ApiDoc;

#[utoipa::path(
    get,
    path = "/api/printers",
    tag = "Printing",
    responses(
        (status = 200, description = "รายการเครื่องพิมพ์ทั้งหมดที่พร้อมใช้งาน", body = PrinterListResponse)
    )
)]
#[get("/api/printers")]
async fn get_printers() -> impl Responder {
    // List of common virtual/software printer keywords to filter out
    let virtual_printer_keywords = [
        "microsoft print to pdf",
        "microsoft xps",
        "onenote",
        "fax",
        "send to onenote",
        "adobe pdf",
        "foxit",
        "nitro pdf",
        "cutepdf",
        "pdf creator",
        "pdfwriter",
        "novapdf",
        "bullzip",
        "dopdf",
        "primopdf",
        "pdf24",
        "microsoft to pdf",
        "send to microsoft",
        "xps document writer",
        "note writer",
    ];

    // ป้องกัน panic โดยใช้ std::panic::catch_unwind
    let printers_result = std::panic::catch_unwind(|| {
        printers::get_printers()
            .into_iter()
            .map(|p| p.name)
            .filter(|name| {
                let name_lower = name.to_lowercase();
                // Filter out printers that contain any of the virtual printer keywords
                !virtual_printer_keywords.iter().any(|keyword| name_lower.contains(keyword))
            })
            .collect::<Vec<String>>()
    });

    let printers = match printers_result {
        Ok(printers_list) => printers_list,
        Err(e) => {
            eprintln!("Error getting printers (caught panic): {:?}", e);
            Vec::new() // Return empty list instead of crashing
        }
    };

    HttpResponse::Ok().json(PrinterListResponse {
        status: "success".to_string(),
        printers,
    })
}

#[utoipa::path(
    get,
    path = "/",
    responses(
        (status = 200, description = "Service status", body = String, content_type = "text/html")
    )
)]
#[get("/")]
async fn index() -> HttpResponse {
    let html_content = r#"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rust Print API Status</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            padding: 20px;
            background: linear-gradient(135deg, #1e1e2e 0%, #282c34 100%);
            color: #ffffff;
            text-align: center;
        }
        .container {
            max-width: 1200px;
            width: 100%;
        }
        .status-message {
            font-size: 2em;
            margin-bottom: 20px;
            color: #61dafb;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.5);
        }
        .logo-container {
            margin-bottom: 40px;
        }
        .logo-text {
            font-family: 'Fira Code', monospace;
            font-size: 5em;
            font-weight: bold;
            color: #dea584;
            text-shadow: 4px 4px 8px rgba(0, 0, 0, 0.7);
            animation: pulse 2s infinite alternate;
        }
        .sub-logo-text {
            font-size: 1.8em;
            color: #f0db4f;
            margin-top: -10px;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.5);
        }
        @keyframes pulse {
            from { transform: scale(1); }
            to { transform: scale(1.05); }
        }
        .printers-section {
            margin-top: 40px;
            width: 100%;
        }
        .section-title {
            font-size: 1.8em;
            color: #61dafb;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }
        .printer-icon {
            font-size: 1.2em;
        }
        .printers-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        .printer-card {
            background: rgba(255, 255, 255, 0.05);
            border: 2px solid rgba(97, 218, 251, 0.3);
            border-radius: 12px;
            padding: 20px;
            transition: all 0.3s ease;
            backdrop-filter: blur(10px);
        }
        .printer-card:hover {
            transform: translateY(-5px);
            border-color: #61dafb;
            box-shadow: 0 8px 20px rgba(97, 218, 251, 0.3);
        }
        .printer-name {
            font-size: 1.1em;
            color: #ffffff;
            word-break: break-word;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
        .loading {
            color: #f0db4f;
            font-size: 1.2em;
            animation: blink 1.5s infinite;
        }
        @keyframes blink {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }
        .error {
            color: #ff6b6b;
            padding: 20px;
            background: rgba(255, 107, 107, 0.1);
            border-radius: 8px;
            border: 1px solid rgba(255, 107, 107, 0.3);
        }
        .no-printers {
            color: #f0db4f;
            padding: 20px;
            background: rgba(240, 219, 79, 0.1);
            border-radius: 8px;
            border: 1px solid rgba(240, 219, 79, 0.3);
        }
        .api-links {
            margin-top: 40px;
            display: flex;
            gap: 20px;
            justify-content: center;
            flex-wrap: wrap;
        }
        .api-link {
            padding: 12px 24px;
            background: rgba(97, 218, 251, 0.1);
            border: 2px solid #61dafb;
            border-radius: 8px;
            color: #61dafb;
            text-decoration: none;
            transition: all 0.3s ease;
            font-weight: 600;
        }
        .api-link:hover {
            background: #61dafb;
            color: #282c34;
            transform: scale(1.05);
        }
        .footer {
            margin-top: 60px;
            padding-top: 20px;
            border-top: 1px solid rgba(255, 255, 255, 0.1);
            color: #8b949e;
            font-size: 0.9em;
            line-height: 1.6;
        }
        .company-name {
            color: #dea584;
            font-weight: 600;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="status-message">🟢 Service Status: Running</div>
        <div class="logo-container">
            <div class="logo-text">Rust</div>
            <div class="sub-logo-text">Print API</div>
        </div>

        <div class="printers-section">
            <h2 class="section-title">
                <span class="printer-icon">🖨️</span>
                Available Printers
            </h2>
            <div id="printers-container">
                <div class="loading">Loading printers...</div>
            </div>
        </div>

        <div class="api-links">
            <a href="/swagger-ui/" class="api-link">📚 API Documentation</a>
            <a href="/api/printers" class="api-link">📋 Printers JSON</a>
        </div>

        <div class="footer">
            <p>Copyright &copy; 2025 <span class="company-name">Bhatara Progress Co., Ltd.</span></p>
            <p>Developed by Bhatara Progress Co., Ltd.</p>
        </div>
    </div>

    <script>
        async function loadPrinters() {
            const container = document.getElementById('printers-container');
            try {
                const response = await fetch('/api/printers');
                const data = await response.json();
                
                if (data.status === 'success' && data.printers && data.printers.length > 0) {
                    const grid = document.createElement('div');
                    grid.className = 'printers-grid';
                    
                    data.printers.forEach(printer => {
                        const card = document.createElement('div');
                        card.className = 'printer-card';
                        card.innerHTML = `
                            <div class="printer-name">
                                <span>🖨️</span>
                                <span>${escapeHtml(printer)}</span>
                            </div>
                        `;
                        grid.appendChild(card);
                    });
                    
                    container.innerHTML = '';
                    container.appendChild(grid);
                } else {
                    container.innerHTML = '<div class="no-printers">No printers found on this system</div>';
                }
            } catch (error) {
                container.innerHTML = `<div class="error">Failed to load printers: ${escapeHtml(error.message)}</div>`;
            }
        }

        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        // Load printers when page loads
        loadPrinters();
    </script>
</body>
</html>
    "#.to_string();
    HttpResponse::Ok().content_type("text/html").body(html_content)
}

#[utoipa::path(
    post,
    path = "/api/print",
    tag = "Printing",
    request_body = PrintRequest,
    responses(
        (status = 200, description = "แปลงและส่งคำสั่งพิมพ์สำเร็จ", body = ResponseMessage),
        (status = 400, description = "เกิดข้อผิดพลาดในการจัดการไฟล์", body = ResponseMessage),
        (status = 500, description = "เกิดข้อผิดพลาดในการประมวลผลหรือสั่งพิมพ์", body = ResponseMessage)
    )
)]
#[post("/api/print")]
async fn print_file_handler(req: web::Json<PrintRequest>) -> impl Responder {
    let base_dir = Path::new("./printable_files");
    let original_file_path = base_dir.join(&req.filename);

    // --- โค้ดที่เปลี่ยน: สร้างชื่อไฟล์ A6 ถาวร โดยมี _a6 ต่อท้าย ---
    let original_filename = &req.filename;
    let a6_filename = original_filename.rfind('.').map_or_else(
        || format!("{}_a6", original_filename), // กรณีไม่มีนามสกุล
        |i| {
            let (name, ext) = original_filename.split_at(i);
            format!("{}_a6{}", name, ext) // กรณีมีนามสกุล เช่น "invoice.pdf" -> "invoice_a6.pdf"
        },
    );
    let a6_file_path = base_dir.join(&a6_filename);
    // ---------------------------------------------------------------------

    if !original_file_path.exists() {
        return HttpResponse::BadRequest().json(ResponseMessage {
            status: "error".to_string(),
            message: format!("File not found: {}", req.filename),
        });
    }

    // 1. แปลงขนาด PDF เป็น A6 และบันทึกไฟล์ใหม่
    match resize_pdf_to_a6(&original_file_path, &a6_file_path) {
        Ok(_) => println!("PDF successfully resized and saved as {}", a6_filename),
        Err(e) => {
            eprintln!("Error resizing PDF: {:?}", e);
            return HttpResponse::InternalServerError().json(ResponseMessage {
                status: "error".to_string(),
                message: format!("Failed to resize PDF to A6: {}", e),
            });
        }
    }

    // 2. อ่านไฟล์ A6 ที่สร้างขึ้นใหม่ และสั่งพิมพ์
    let file_data = match std::fs::read(&a6_file_path) {
        Ok(data) => data,
        Err(e) => {
            eprintln!("Error reading A6 file {}: {:?}", a6_filename, e);
            return HttpResponse::InternalServerError().json(ResponseMessage {
                status: "error".to_string(),
                message: format!("Failed to read A6 file {}. Error: {}", a6_filename, e),
            });
        }
    };

    println!("Successfully read A6 file: {}", a6_filename);

    let printer = match printers::get_printer_by_name(&req.printer_name) {
        Some(p) => p,
        None => {
            return HttpResponse::BadRequest().json(ResponseMessage {
                status: "error".to_string(),
                message: format!("Printer not found: {}", req.printer_name),
            });
        }
    };

    let options = PrinterJobOptions {
        name: Some(&format!("A6 Print Job - {}", req.filename)),
        raw_properties: &[],
    };

    match printer.print(&file_data, options) {
        Ok(_) => {
            println!("Print job sent successfully to {}", req.printer_name);
            HttpResponse::Ok().json(ResponseMessage {
                status: "success".to_string(),
                message: format!(
                    "Resized to A6, saved as {}, and sent to printer {}",
                    a6_filename, req.printer_name
                ),
            })
        }
        Err(e) => {
            eprintln!("Error sending print job: {:?}", e);
            HttpResponse::InternalServerError().json(ResponseMessage {
                status: "error".to_string(),
                message: format!("Failed to send print job: {:?}", e),
            })
        }
    }
}

async fn run_app() -> std::io::Result<()> {
    let base_dir = Path::new("./printable_files");
    if !base_dir.exists() {
        std::fs::create_dir(base_dir)?;
        println!("Created directory: ./printable_files");
    }

    let openapi = web::Data::new(ApiDoc::openapi());

    println!("Starting server at http://127.0.0.1:3000");
    println!("Swagger UI available at: http://127.0.0.1:3000/swagger-ui/");

    HttpServer::new(move || {
        App::new()
            .app_data(openapi.clone())
            .service(index)
            .service(get_printers)
            .service(print_file_handler)
            .service(
                SwaggerUi::new("/swagger-ui/{_:.*}")
                    .url("/api-docs/openapi.json", openapi.get_ref().clone()),
            )
    })
    .bind(("127.0.0.1", 3000))?
    .run()
    .await
}

fn my_service_main(_arguments: Vec<OsString>) {
    if let Err(e) = run_service() {
        eprintln!("Failed to run service: {}", e);
    }
}

fn run_service() -> windows_service::Result<()> {
    let (shutdown_tx, shutdown_rx) = mpsc::channel();

    let event_handler = move |control_event| -> ServiceControlHandlerResult {
        match control_event {
            ServiceControl::Stop => {
                shutdown_tx.send(()).unwrap();
                ServiceControlHandlerResult::NoError
            }
            ServiceControl::Interrogate => ServiceControlHandlerResult::NoError,
            _ => ServiceControlHandlerResult::NotImplemented,
        }
    };

    let status_handle = service_control_handler::register(SERVICE_NAME, event_handler)?;

    status_handle.set_service_status(ServiceStatus {
        service_type: SERVICE_TYPE,
        current_state: ServiceState::Running,
        controls_accepted: ServiceControlAccept::STOP,
        exit_code: ServiceExitCode::Win32(0),
        checkpoint: 0,
        wait_hint: Duration::default(),
        process_id: None,
    })?;

    let rt = tokio::runtime::Runtime::new().unwrap();
    std::thread::spawn(move || {
        rt.block_on(async {
            if let Err(e) = run_app().await {
                eprintln!("Server failed to start: {}", e);
            }
        });
    });

    loop {
        match shutdown_rx.recv_timeout(Duration::from_secs(1)) {
            Ok(_) => break,
            Err(mpsc::RecvTimeoutError::Timeout) => (),
            Err(_) => break,
        }
    }

    status_handle.set_service_status(ServiceStatus {
        service_type: SERVICE_TYPE,
        current_state: ServiceState::Stopped,
        controls_accepted: ServiceControlAccept::empty(),
        exit_code: ServiceExitCode::Win32(0),
        checkpoint: 0,
        wait_hint: Duration::default(),
        process_id: None,
    })?;

    Ok(())
}

define_windows_service!(ffi_service_main, my_service_main);

fn main() -> windows_service::Result<()> {
    let args: Vec<String> = std::env::args().collect();
    if args.len() > 1 && args[1] == "--console" {
        // Run in console mode
        if let Err(e) = tokio::runtime::Runtime::new().unwrap().block_on(run_app()) {
            eprintln!("Failed to run in console mode: {}", e);
        }
    } else {
        // Run as a service
        service_dispatcher::start(SERVICE_NAME, ffi_service_main)?;
    }
    Ok(())
}
