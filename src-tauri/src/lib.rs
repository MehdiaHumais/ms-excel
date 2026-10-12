use std::fs;
use std::path::Path;
use std::sync::Mutex;

use flate2::read::GzDecoder;
use flate2::write::GzEncoder;
use flate2::Compression;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use tauri::{Emitter, Manager};

/// Native BritSheet container: 10-byte magic + version byte + gzip(JSON).
const MAGIC: &[u8; 10] = b"BRITSHEET1";
const VERSION: u8 = 1;

/// Event the shell emits to hand a workbook path to the webview.
const OPEN_FILE_EVENT: &str = "britsheet://open-file";

/// Extensions the frontend loader knows how to open.
const OPENABLE_EXTS: [&str; 7] = ["britsheet", "bs", "xlsx", "xls", "csv", "txt", "json"];

pub struct DbState(Mutex<rusqlite::Connection>);

/// Workbook path this process was launched with, set once at startup.
pub struct LaunchState(Mutex<Option<String>>);

/// Pick the first workbook path out of a process command line.
///
/// Windows launches an associated file by appending its path to the target
/// command, so `britsheet.exe C:\...\Budget.britsheet` must open that file.
/// Flags and directories are ignored, and the file has to actually exist so a
/// stray argument cannot be mistaken for a document.
fn workbook_from_args(args: &[String]) -> Option<String> {
    args.iter()
        .skip(1)
        .filter(|a| !a.starts_with('-'))
        .map(|a| a.trim().trim_matches('"'))
        .find(|a| {
            let lower = a.to_ascii_lowercase();
            OPENABLE_EXTS
                .iter()
                .any(|ext| lower.ends_with(&format!(".{ext}")))
                && Path::new(a).is_file()
        })
        .map(|a| a.to_string())
}

/// The workbook this instance was launched with, if any.
#[tauri::command]
fn launch_file(state: tauri::State<'_, LaunchState>) -> Option<String> {
    state.0.lock().ok().and_then(|g| g.clone())
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RecentFile {
    pub path: String,
    pub name: String,
    pub modified: i64,
}

fn open_db(app_data: &Path) -> Result<rusqlite::Connection, String> {
    fs::create_dir_all(app_data).map_err(|e| format!("create data dir: {e}"))?;
    let conn = rusqlite::Connection::open(app_data.join("britsheet.db")).map_err(|e| e.to_string())?;
    conn.execute_batch(
        "CREATE TABLE IF NOT EXISTS recent (
            path TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            modified INTEGER NOT NULL
        );",
    )
    .map_err(|e| e.to_string())?;
    Ok(conn)
}

fn gzip_json(json: &str) -> Result<Vec<u8>, String> {
    let mut enc = GzEncoder::new(Vec::new(), Compression::default());
    std::io::Write::write_all(&mut enc, json.as_bytes()).map_err(|e| e.to_string())?;
    enc.finish().map_err(|e| e.to_string())
}

fn gunzip_json(bytes: &[u8]) -> Result<String, String> {
    let mut dec = GzDecoder::new(bytes);
    let mut out = String::new();
    std::io::Read::read_to_string(&mut dec, &mut out).map_err(|e| e.to_string())?;
    Ok(out)
}

fn encode_workbook(json: &str) -> Result<Vec<u8>, String> {
    let mut out = Vec::with_capacity(MAGIC.len() + 1 + json.len());
    out.extend_from_slice(MAGIC);
    out.push(VERSION);
    out.extend_from_slice(&gzip_json(json)?);
    Ok(out)
}

fn decode_workbook(bytes: &[u8]) -> Result<String, String> {
    if bytes.len() < MAGIC.len() + 1 {
        return Err("Not a BritSheet file (too short)".into());
    }
    if &bytes[..MAGIC.len()] != MAGIC {
        return Err("Not a BritSheet file (bad magic header)".into());
    }
    let version = bytes[MAGIC.len()];
    if version != VERSION {
        return Err(format!("Unsupported BritSheet version {version}"));
    }
    gunzip_json(&bytes[MAGIC.len() + 1..])
}

/// Atomically write a BritSheet file: write to temp sibling, then rename.
fn atomic_write(path: &Path, data: &[u8]) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("create parent dir: {e}"))?;
    }
    let tmp = path.with_extension("britsheet.tmp");
    fs::write(&tmp, data).map_err(|e| format!("write temp: {e}"))?;
    fs::rename(&tmp, path).map_err(|e| {
        let _ = fs::remove_file(&tmp);
        format!("commit rename: {e}")
    })
}

#[tauri::command]
fn save_workbook(path: String, json: String) -> Result<(), String> {
    let data = encode_workbook(&json)?;
    atomic_write(Path::new(&path), &data)
}

#[tauri::command]
fn open_workbook(path: String) -> Result<String, String> {
    let bytes = fs::read(&path).map_err(|e| format!("read: {e}"))?;
    decode_workbook(&bytes)
}

#[tauri::command]
fn read_text_file(path: String) -> Result<String, String> {
    let bytes = fs::read(&path).map_err(|e| e.to_string())?;
    let text = String::from_utf8(bytes).map_err(|e| e.to_string())?;
    Ok(text)
}

#[tauri::command]
fn read_binary_file(path: String) -> Result<String, String> {
    let bytes = fs::read(&path).map_err(|e| e.to_string())?;
    Ok(base64_encode(&bytes))
}

fn base64_encode(bytes: &[u8]) -> String {
    const TABLE: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((bytes.len() + 2) / 3 * 4);
    let mut i = 0;
    while i + 3 <= bytes.len() {
        let n = (bytes[i] as u32) << 16 | (bytes[i + 1] as u32) << 8 | bytes[i + 2] as u32;
        out.push(TABLE[(n >> 18) as usize] as char);
        out.push(TABLE[((n >> 12) & 63) as usize] as char);
        out.push(TABLE[((n >> 6) & 63) as usize] as char);
        out.push(TABLE[(n & 63) as usize] as char);
        i += 3;
    }
    let rem = bytes.len() - i;
    if rem == 1 {
        let n = (bytes[i] as u32) << 16;
        out.push(TABLE[(n >> 18) as usize] as char);
        out.push(TABLE[((n >> 12) & 63) as usize] as char);
        out.push('=');
        out.push('=');
    } else if rem == 2 {
        let n = (bytes[i] as u32) << 16 | (bytes[i + 1] as u32) << 8;
        out.push(TABLE[(n >> 18) as usize] as char);
        out.push(TABLE[((n >> 12) & 63) as usize] as char);
        out.push(TABLE[((n >> 6) & 63) as usize] as char);
        out.push('=');
    }
    out
}

#[tauri::command]
fn list_recent(state: tauri::State<DbState>) -> Result<Vec<RecentFile>, String> {
    let conn = state.0.lock().map_err(|e| e.to_string())?;
    let mut stmt = conn
        .prepare("SELECT path, name, modified FROM recent ORDER BY modified DESC LIMIT 12")
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map([], |r| {
            Ok(RecentFile {
                path: r.get(0)?,
                name: r.get(1)?,
                modified: r.get(2)?,
            })
        })
        .map_err(|e| e.to_string())?;
    let mut files = Vec::new();
    for row in rows {
        files.push(row.map_err(|e| e.to_string())?);
    }
    Ok(files)
}

#[tauri::command]
fn add_recent(state: tauri::State<DbState>, path: String) -> Result<(), String> {
    let conn = state.0.lock().map_err(|e| e.to_string())?;
    let name = Path::new(&path)
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| path.clone());
    let modified = chrono_now_ms() as i64;
    conn.execute(
        "INSERT OR REPLACE INTO recent (path, name, modified) VALUES (?1, ?2, ?3)",
        rusqlite::params![path, name, modified],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn remove_recent(state: tauri::State<DbState>, path: String) -> Result<(), String> {
    let conn = state.0.lock().map_err(|e| e.to_string())?;
    conn.execute("DELETE FROM recent WHERE path = ?1", rusqlite::params![path])
        .map_err(|e| e.to_string())?;
    Ok(())
}

fn chrono_now_ms() -> u128 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0)
}

#[tauri::command]
fn default_data_path(app: tauri::AppHandle) -> Result<String, String> {
    let dir = app
        .path()
        .app_data_dir()
        .map_err(|e| e.to_string())?;
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.to_string_lossy().to_string())
}

/// Validate that a string parses as JSON object (sanity check before save).
#[tauri::command]
fn validate_json_object(json: String) -> Result<bool, String> {
    match serde_json::from_str::<Value>(&json) {
        Ok(Value::Object(_)) => Ok(true),
        Ok(_) => Ok(false),
        Err(e) => Err(format!("invalid JSON: {e}")),
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_single_instance::init(|app, args, _cwd| {
            // A second launch (double-clicking a workbook in Explorer) is
            // routed here by the single-instance plugin. Forward the file to
            // the window that is already open instead of dropping it.
            if let Some(path) = workbook_from_args(&args) {
                let _ = app.emit(OPEN_FILE_EVENT, path);
            }
            if let Some(win) = app.get_webview_window("main") {
                let _ = win.unminimize();
                let _ = win.show();
                let _ = win.set_focus();
            }
        }))
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            let dir = app.path().app_data_dir()?;
            let conn = open_db(&dir)?;
            app.manage(DbState(Mutex::new(conn)));
            let args: Vec<String> = std::env::args().collect();
            app.manage(LaunchState(Mutex::new(workbook_from_args(&args))));
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            save_workbook,
            open_workbook,
            read_text_file,
            read_binary_file,
            list_recent,
            add_recent,
            remove_recent,
            default_data_path,
            validate_json_object,
            launch_file
        ])
        .run(tauri::generate_context!())
        .expect("error while running BritSheet");
}