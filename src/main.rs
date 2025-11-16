use actix_web::{post, web, App, HttpResponse, HttpServer, Responder};
use serde::{Deserialize, Serialize};
use serde_json::json;
use std::path::{Path, PathBuf};
use printers::{self, common::base::job::PrinterJobOptions};
use anyhow::{Result, Context, bail}; 
use lopdf::content::{Content, Operation};
use lopdf::Document;
use lopdf::Object;
use utoipa::{OpenApi, ToSchema};
use utoipa_swagger_ui::SwaggerUi;

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

        let pages_to_process: Vec<(lopdf::ObjectId, lopdf::ObjectId)> = doc.get_pages()
            .values()
            .filter_map(|page_id| {
                doc.get_dictionary(*page_id).ok().and_then(|page| {
                    if let Ok(Object::Reference(content_ref)) = page.get(b"Contents") {
                        Some((*page_id, *content_ref))
                    } else {
                        None
                    }
                })
            })
            .collect();
    
        for (page_id, content_ref) in pages_to_process {
            // Modify MediaBox
            if let Ok(page) = doc.get_dictionary_mut(page_id) {
                let new_media_box = vec![
                    Object::Real(0.0), Object::Real(0.0),
                    Object::Real(A6_WIDTH_PTS), Object::Real(A6_HEIGHT_PTS),
                ];
                page.set("MediaBox", Object::Array(new_media_box));
            }
    
            // Modify content stream
            if let Ok(content_stream) = doc.get_object(content_ref).and_then(|obj| obj.as_stream()) {
                let mut content = Content::decode(&content_stream.content)?;
    
                let matrix_op = Operation::new("cm", vec![
                    Object::Real(scale_factor),
                    Object::Real(0.0),
                    Object::Real(0.0),
                    Object::Real(scale_factor),
                    Object::Real(0.0),
                    Object::Real(0.0),
                ]);
                content.operations.insert(0, matrix_op);
    
                let new_content_bytes = content.encode()?;
                let new_stream = lopdf::Stream {
                    dict: content_stream.dict.clone(),
                    content: new_content_bytes,
                    allows_compression: content_stream.allows_compression,
                    start_position: content_stream.start_position,
                };
                doc.objects.insert(content_ref, Object::Stream(new_stream));
            }
        }
    // 4. บันทึกเอกสาร A6 ใหม่
    doc.prune_objects();
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
    paths(print_file_handler),
    components(schemas(PrintRequest, ResponseMessage)),
    tags((name = "Printing", description = "Endpoints สำหรับการดำเนินการสั่งพิมพ์ไฟล์และแปลงขนาด"))
)]
struct ApiDoc;


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
        }
    );
    let a6_file_path = base_dir.join(&a6_filename);
    // ---------------------------------------------------------------------

    if !original_file_path.exists() {
        return HttpResponse::BadRequest().json(
             ResponseMessage { 
                status: "error".to_string(), 
                message: format!("File not found: {}", req.filename) 
            }
        );
    }
    
    // 1. แปลงขนาด PDF เป็น A6 และบันทึกไฟล์ใหม่
    match resize_pdf_to_a6(&original_file_path, &a6_file_path) {
        Ok(_) => println!("PDF successfully resized and saved as {}", a6_filename),
        Err(e) => {
            eprintln!("Error resizing PDF: {:?}", e);
            return HttpResponse::InternalServerError().json(
                 ResponseMessage { 
                    status: "error".to_string(), 
                    message: format!("Failed to resize PDF to A6: {}", e) 
                }
            );
        }
    }

    // 2. อ่านไฟล์ A6 ที่สร้างขึ้นใหม่ และสั่งพิมพ์
    // 2. อ่านไฟล์ A6 ที่สร้างขึ้นใหม่ และสั่งพิมพ์
    let file_data = match std::fs::read(&a6_file_path) {
        Ok(data) => data,
        Err(e) => {
            eprintln!("Error reading A6 file {}: {:?}", a6_filename, e);
            return HttpResponse::InternalServerError().json(
                 ResponseMessage {
                    status: "error".to_string(),
                    message: format!("Failed to read A6 file {}. Error: {}", a6_filename, e)
                }
            );
        }
    };

    println!("Successfully read A6 file: {}", a6_filename);

    let printer = match printers::get_printer_by_name(&req.printer_name) {
        Some(p) => p,
        None => {
            return HttpResponse::BadRequest().json(
                 ResponseMessage {
                    status: "error".to_string(),
                    message: format!("Printer not found: {}", req.printer_name)
                }
            );
        }
    };

                            let options = PrinterJobOptions {

                                name: Some(&format!("A6 Print Job - {}", req.filename)),

                                raw_properties: &[],

                            };

                match printer.print(&file_data, options) {
        Ok(_) => {
            println!("Print job sent successfully to {}", req.printer_name);
            HttpResponse::Ok().json(
                 ResponseMessage {
                    status: "success".to_string(),
                    message: format!("Resized to A6, saved as {}, and sent to printer {}", a6_filename, req.printer_name)
                }
            )
        }
        Err(e) => {
            eprintln!("Error sending print job: {:?}", e);
            HttpResponse::InternalServerError().json(
                 ResponseMessage {
                    status: "error".to_string(),
                    message: format!("Failed to send print job: {:?}", e)
                }
            )
        }
    }
}


#[tokio::main]
async fn main() -> std::io::Result<()> {
    let base_dir = Path::new("./printable_files");
    if !base_dir.exists() {
        std::fs::create_dir(base_dir)?;
        println!("Created directory: ./printable_files");
    }

    let openapi = ApiDoc::openapi();

    println!("Starting server at http://127.0.0.1:8080");
    println!("Swagger UI available at: http://127.0.0.1:8080/swagger-ui/");

    HttpServer::new(move || {
        App::new()
            .service(print_file_handler)
            .service(
                SwaggerUi::new("/swagger-ui/{_:.*}")
                    .url("/api-docs/openapi.json", openapi.clone())
            )
    })
    .bind(("127.0.0.1", 8080))?
    .run()
    .await
}