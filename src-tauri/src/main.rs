#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::process::Command;
#[tauri::command]
fn extract_docx(path: String) -> Result<String, String> {
    let mut script_path = std::env::current_dir()
        .map_err(|e| format!("Failed to resolve current dir: {e}"))?;
    script_path.push("../tools/docx_extract.py");
    let script_path = script_path
        .canonicalize()
        .map_err(|e| format!("Failed to resolve extractor path: {e}"))?;

    let output = Command::new("python3")
        .arg(script_path)
        .arg(path)
        .output()
        .map_err(|e| format!("Failed to start python: {e}"))?;

    if !output.status.success() {
        let err = String::from_utf8_lossy(&output.stderr).to_string();
        return Err(err);
    }

    let stdout = String::from_utf8_lossy(&output.stdout).to_string();
    Ok(stdout)
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![extract_docx])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
