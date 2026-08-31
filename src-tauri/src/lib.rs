use base64::Engine;
use reqwest::header::{HeaderMap, HeaderValue, ACCEPT, ACCEPT_LANGUAGE, COOKIE, REFERER, USER_AGENT};
use serde::Serialize;

const REPORT_BASE_URL: &str = "https://applinde.a365.com.pe/reportes/reporte_matriz";

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct MatrixReport {
    data_base64: String,
    content_type: String,
    file_name: String,
}

fn is_valid_date(value: &str) -> bool {
    value.len() == 10
        && value.chars().enumerate().all(|(index, current)| match index {
            4 | 7 => current == '-',
            _ => current.is_ascii_digit(),
        })
}

fn is_valid_session(value: &str) -> bool {
    !value.is_empty() && value.chars().all(|current| current.is_ascii_alphanumeric())
}

fn normalize_session(value: &str) -> String {
    value
        .split(';')
        .find_map(|part| part.trim().strip_prefix("PHPSESSID="))
        .unwrap_or_else(|| value.trim().strip_prefix("PHPSESSID=").unwrap_or(value.trim()))
        .trim_matches(|current| current == '"' || current == '\'')
        .trim()
        .to_string()
}

fn looks_like_login_page(report: &[u8], content_type: &str) -> bool {
    if !content_type.to_lowercase().contains("text/html") {
        return false;
    }

    let preview = String::from_utf8_lossy(&report[..report.len().min(5000)]).to_lowercase();
    let has_report_table = preview.contains("<table")
        && (preview.contains("subject_email")
            || preview.contains("date_email")
            || preview.contains("fecha_asignacion")
            || preview.contains("usuario asignado"));

    !has_report_table
        && (preview.contains("login")
            || preview.contains("password")
            || preview.contains("contraseña")
            || preview.contains("iniciar ses")
            || preview.contains("phpsessid"))
}

fn report_headers(session_id: &str) -> Result<HeaderMap, String> {
    let mut headers = HeaderMap::new();
    headers.insert(ACCEPT, HeaderValue::from_static("text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;q=0.8,*/*;q=0.7"));
    headers.insert(ACCEPT_LANGUAGE, HeaderValue::from_static("es-419,es;q=0.9"));
    headers.insert(REFERER, HeaderValue::from_static("https://applinde.a365.com.pe/reportes/"));
    headers.insert(USER_AGENT, HeaderValue::from_static("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36"));
    headers.insert(
        COOKIE,
        HeaderValue::from_str(&format!("PHPSESSID={session_id}"))
            .map_err(|_| "PHPSESSID invalido.".to_string())?,
    );
    headers.insert("sec-fetch-dest", HeaderValue::from_static("document"));
    headers.insert("sec-fetch-mode", HeaderValue::from_static("navigate"));
    headers.insert("sec-fetch-site", HeaderValue::from_static("same-origin"));
    headers.insert("upgrade-insecure-requests", HeaderValue::from_static("1"));
    Ok(headers)
}

#[allow(non_snake_case)]
#[tauri::command]
async fn download_matrix_report(
    startDate: String,
    endDate: String,
    sessionId: String,
) -> Result<MatrixReport, String> {
    let start_date = startDate.trim();
    let end_date = endDate.trim();
    let session_id = normalize_session(&sessionId);

    if !is_valid_date(start_date) || !is_valid_date(end_date) {
        return Err("Selecciona fechas validas.".to_string());
    }

    if !is_valid_session(&session_id) {
        return Err("Ingresa un PHPSESSID valido.".to_string());
    }

    let url = format!("{REPORT_BASE_URL}/{start_date}/{end_date}/");
    let client = reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::limited(10))
        .build()
        .map_err(|_| "No pude preparar la conexion local.".to_string())?;
    let response = client
        .get(url)
        .headers(report_headers(&session_id)?)
        .send()
        .await
        .map_err(|_| "No pude conectarme con A365 desde esta PC.".to_string())?;

    if !response.status().is_success() {
        return Err(format!("No pude descargar el reporte. Codigo {}.", response.status()));
    }

    let content_type = response
        .headers()
        .get(reqwest::header::CONTENT_TYPE)
        .and_then(|value| value.to_str().ok())
        .unwrap_or("application/octet-stream")
        .to_string();
    let bytes = response
        .bytes()
        .await
        .map_err(|_| "No pude leer la respuesta de A365.".to_string())?;

    if looks_like_login_page(&bytes, &content_type) {
        return Err("A365 devolvio login o sesion expirada. Vuelve a copiar el PHPSESSID.".to_string());
    }

    Ok(MatrixReport {
        data_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        content_type,
        file_name: format!("reporte-matriz-{start_date}-{end_date}.xlsx"),
    })
}

pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![download_matrix_report])
        .run(tauri::generate_context!())
        .expect("error while running A365 desktop app");
}
