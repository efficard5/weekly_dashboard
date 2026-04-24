import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import os
import json
import mimetypes
import streamlit.components.v1 as components


try:
    from streamlit.errors import StreamlitSecretNotFoundError
except ImportError:
    StreamlitSecretNotFoundError = Exception

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
except ImportError:
    service_account = None
    build = None
    MediaIoBaseUpload = None
    MediaIoBaseDownload = None



# --- PAGE SETUP ---
st.set_page_config(page_title="Industrial Automation PMO", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border: 1px solid #e0e0e0; }
    .card { background-color: #ffffff; padding: 20px; border-radius: 8px; border: 1px dotted #ccc; margin-bottom: 10px; }
    .sidebar .sidebar-content { background-image: linear-gradient(#2e3b4e, #2e3b4e); color: white; }
    .topic-image-container { border-radius: 8px; overflow: hidden; margin-bottom: 6px; }
    </style>
    """, unsafe_allow_html=True)

# Inject Auto-Bullet Javascript for textareas
components.html(
    """
    <script>
    const doc = window.parent.document;
    if (!doc.getElementById("auto-bullet-script")) {
        const script = doc.createElement("script");
        script.id = "auto-bullet-script";
        script.innerHTML = `
            function pushValueToStreamlit(textarea, value, caretPos) {
                let nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                nativeInputValueSetter.call(textarea, value);
                textarea.selectionStart = textarea.selectionEnd = caretPos;
                textarea.dispatchEvent(new Event("input", { bubbles: true }));
            }

            document.addEventListener("keydown", function(e) {
                if (e.target.tagName === "TEXTAREA" && e.key === "Enter") {
                    const val = e.target.value;
                    const start = e.target.selectionStart;
                    const end = e.target.selectionEnd;
                    
                    const textBeforeCursor = val.substring(0, start);
                    const lastNewLine = textBeforeCursor.lastIndexOf("\\n");
                    const currentLine = textBeforeCursor.substring(lastNewLine + 1);
                    
                    const bulletMatch = currentLine.match(/^(\\s*[-*•]\\s+)/);
                    if (bulletMatch) {
                        if (currentLine.trim() === bulletMatch[1].trim()) {
                            e.preventDefault();
                            const lineStart = lastNewLine !== -1 ? lastNewLine + 1 : 0;
                            const updatedValue = val.substring(0, lineStart) + val.substring(end);
                            pushValueToStreamlit(e.target, updatedValue, lineStart);
                            return;
                        }

                        e.preventDefault();
                        const bullet = bulletMatch[1];
                        const newText = "\\n" + bullet;
                        const updatedValue = val.substring(0, start) + newText + val.substring(end);
                        pushValueToStreamlit(e.target, updatedValue, start + newText.length);
                        return;
                    }

                    if (currentLine.trim() !== "") {
                        e.preventDefault();
                        const newText = "\\n- ";
                        const updatedValue = val.substring(0, start) + newText + val.substring(end);
                        pushValueToStreamlit(e.target, updatedValue, start + newText.length);
                    }
                }
            });
        `;
        doc.body.appendChild(script);
    }
    </script>
    """,
    height=0,
    width=0
)

# --- EXCEL DATA LAYER ---
@st.cache_data
def load_data():
    file_path = "data/tasks.xlsx"
    required_cols = [
        "Project", "Topic", "Task Name", "Start Date", "End Date", 
        "Completion %", "Status", "Employee", "Week", "Hidden",
        "Milestone_Text", "Milestone_Role", "Milestone_Author_Name"
    ]
    
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        needs_save = False
        # Consolidate column defaults and type enforcement into a single loop
        defaults = {
            "Hidden": False, "Completion %": 0, "Week": 1, 
            "Project": "R&D Project", "Employee": "Unassigned",
            "Milestone_Role": "None"
        }
        for col in required_cols:
            if col not in df.columns:
                df[col] = defaults.get(col, "")
                needs_save = True
        
        df["Hidden"] = df["Hidden"].fillna(False).astype(bool)
        if needs_save:
            os.makedirs("data", exist_ok=True)
            df.to_excel(file_path, index=False)
        return df
    return pd.DataFrame(columns=required_cols)

def save_data(data_df):
    os.makedirs("data", exist_ok=True)
    data_df.to_excel("data/tasks.xlsx", index=False)
    sync_data_file_to_drive("data/tasks.xlsx")
    log_activity("Updated tasks.xlsx")
    st.cache_data.clear()  # Ensure the next load reflects the newly saved data

def load_notes():
    if os.path.exists("data/project_notes.json"):
        with open("data/project_notes.json", "r") as f:
            return json.load(f)
    return {}

def save_notes(notes):
    os.makedirs("data", exist_ok=True)
    with open("data/project_notes.json", "w") as f:
        json.dump(notes, f, indent=4)
    sync_data_file_to_drive("data/project_notes.json")
    log_activity("Updated project_notes.json")

def load_planned_milestones():
    if os.path.exists("data/planned_milestones.json"):
        with open("data/planned_milestones.json", "r") as f:
            return json.load(f)
    return {}

def save_planned_milestones(milestones):
    os.makedirs("data", exist_ok=True)
    with open("data/planned_milestones.json", "w") as f:
        json.dump(milestones, f, indent=4)
    sync_data_file_to_drive("data/planned_milestones.json")
    log_activity("Updated planned_milestones.json")

def load_drive_metadata():
    if os.path.exists("data/drive_metadata.json"):
        with open("data/drive_metadata.json", "r") as f:
            return json.load(f)
    return {}

def save_drive_metadata(data):
    os.makedirs("data", exist_ok=True)
    with open("data/drive_metadata.json", "w") as f:
        json.dump(data, f, indent=4)
    sync_data_file_to_drive("data/drive_metadata.json")

def load_competitor_data():
    file_path = "data/competitors.xlsx"
    if os.path.exists(file_path):
        try:
            xls = pd.read_excel(file_path, sheet_name=None)
            safe_data = {}
            for sheet, df in xls.items():
                if sheet == "Empty" and df.empty:
                    continue
                if df.empty:
                    cols = df.columns.tolist()
                    safe_data[sheet] = [{c: "" for c in cols}]
                else:
                    safe_data[sheet] = df.fillna("").to_dict(orient="records")
            if safe_data:
                return safe_data
        except Exception:
            pass

    return {}

def save_competitor_data(data):
    os.makedirs("data", exist_ok=True)
    file_path = "data/competitors.xlsx"
    import pandas as pd
    try:
        with pd.ExcelWriter(file_path) as writer:
            if not data:
                pd.DataFrame(columns=["Competitor", "Value", "Notes"]).to_excel(writer, sheet_name="Empty", index=False)
            else:
                for topic, rows in data.items():
                    safe_sheet = str(topic).replace("/", "_").replace("\\", "_").replace("?", "").replace("*", "").replace("[", "").replace("]", "")[:31]
                    if not rows:
                        pd.DataFrame(columns=["Competitor", "Value", "Notes"]).to_excel(writer, sheet_name=safe_sheet, index=False)
                    else:
                        pd.DataFrame(rows).to_excel(writer, sheet_name=safe_sheet, index=False)
                        
        sync_data_file_to_drive(file_path)
    except Exception as e:
        import streamlit as st
        st.error(f"Failed to save excel: {e}")

GOOGLE_DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]

def _get_streamlit_secret(key, default=None):
    try:
        return st.secrets[key]
    except (StreamlitSecretNotFoundError, KeyError, FileNotFoundError, AttributeError):
        return default
    except Exception:
        return default

def _load_google_drive_credentials():
    import os
    import json
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
    except ImportError:
        Credentials = None
        Request = None

    creds = None

    if os.path.exists('token.json') and Credentials is not None:
        try:
            creds = Credentials.from_authorized_user_file('token.json', GOOGLE_DRIVE_SCOPES)
        except Exception:
            pass

    if creds and creds.expired and creds.refresh_token:
        try:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        except Exception as e:
            print(f"Token refresh failed: {e}")
            creds = None
            # If token is revoked/invalid, rename it to prevent further attempts
            if os.path.exists('token.json'):
                try:
                    os.rename('token.json', 'token.json.bak')
                except:
                    pass

    if creds and creds.valid:
        return creds

    # Token existed but was invalid/expired and couldn't be refreshed automatically.
    # We should prefer Service Account fallback before trying interactive OAuth.
    
    if service_account is not None:
        # 1. Check file path (optional)
        service_account_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
        if service_account_path and os.path.exists(service_account_path):
            try:
                return service_account.Credentials.from_service_account_file(
                    service_account_path,
                    scopes=GOOGLE_DRIVE_SCOPES
                )
            except:
                pass

        # 2. Check JSON string (optional)
        service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        if service_account_json:
            try:
                return service_account.Credentials.from_service_account_info(
                    json.loads(service_account_json),
                    scopes=GOOGLE_DRIVE_SCOPES
                )
            except Exception:
                pass

        # 3. Read from Streamlit secrets
        try:
            secret_account = st.secrets.get("service_account") or st.secrets.get("gdrive_service_account")
            if secret_account:
                return service_account.Credentials.from_service_account_info(
                    dict(secret_account),
                    scopes=GOOGLE_DRIVE_SCOPES
                )
        except Exception as e:
            try:
                st.sidebar.error(f"Secrets error: {e}")
            except:
                pass

    # If Service Account is not available, try OAuth Flow as a last resort
    if os.path.exists('credentials.json'):
        try:
            from google_auth_oauthlib.flow import InstalledAppFlow
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', GOOGLE_DRIVE_SCOPES)
            creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
            return creds
        except Exception as e:
            try:
                st.sidebar.error(f"OAuth Flow error: {e}")
            except:
                pass

    return None


@st.cache_resource(show_spinner=False)
def get_google_drive_service():
    try:
        from googleapiclient.discovery import build

        credentials = _load_google_drive_credentials()
        if not credentials:
            return None

        return build("drive", "v3", credentials=credentials)

    except Exception as e:
        st.error(f"Google Drive error: {e}")
        return None

def get_google_drive_root_folder_id():
    root_id = os.getenv("GOOGLE_DRIVE_ROOT_FOLDER_ID", "").strip()
    if root_id:
        return root_id
    root_id = os.getenv("gdrive_root_folder_id", "").strip()
    if root_id:
        return root_id
    secret_root_id = _get_streamlit_secret("gdrive_root_folder_id", "")
    if secret_root_id:
        return str(secret_root_id).strip()
    return "root"

def google_drive_is_ready():
    return get_google_drive_service() is not None



def get_google_drive_debug_info():
    root_id = get_google_drive_root_folder_id()
    service = get_google_drive_service()
    credentials = _load_google_drive_credentials()
    return {
        "root_id": root_id,
        "service_ready": service is not None,
        "credentials_ready": credentials is not None,
        "drive_ready": bool(root_id) and service is not None,
        "service_repr": repr(service) if service is not None else "None",
    }

def escape_drive_query_value(value):
    return str(value).replace("\\", "\\\\").replace("'", "\\'")

def ensure_drive_folder(service, folder_name, parent_id):
    escaped_folder_name = escape_drive_query_value(folder_name)
    query = (
        f"name = '{escaped_folder_name}' and "
        f"'{parent_id}' in parents and "
        "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    result = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name)",
        pageSize=1,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    files = result.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": str(folder_name),
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id]
    }
    folder = service.files().create(
        body=metadata, 
        fields="id",
        supportsAllDrives=True
    ).execute()
    return folder["id"]

def ensure_drive_path(path_parts):
    service = get_google_drive_service()
    root_folder_id = get_google_drive_root_folder_id()
    if service is None or not root_folder_id:
        return None

    parent_id = root_folder_id
    for part in path_parts:
        clean_part = str(part or "").strip()
        if clean_part == "":
            continue
        parent_id = ensure_drive_folder(service, clean_part, parent_id)
    return parent_id

def upload_bytes_to_drive(path_parts, file_name, file_bytes, mime_type=None):
    service = get_google_drive_service()
    if service is None:
        raise RuntimeError("Google Drive service is not configured.")

    parent_id = ensure_drive_path(path_parts)
    if parent_id is None:
        raise RuntimeError("Google Drive root folder is not configured.")

    mime_type = mime_type or mimetypes.guess_type(file_name)[0] or "application/octet-stream"
    escaped_file_name = escape_drive_query_value(file_name)
    existing_query = (
        f"name = '{escaped_file_name}' and "
        f"'{parent_id}' in parents and trashed = false"
    )
    existing = service.files().list(
        q=existing_query,
        spaces="drive",
        fields="files(id, name)",
        pageSize=1,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute().get("files", [])

    import io
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type, resumable=False)
    metadata = {"name": file_name, "parents": [parent_id]}
    if existing:
        file_obj = service.files().update(
            fileId=existing[0]["id"],
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True
        ).execute()
    else:
        file_obj = service.files().create(
            body=metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True
        ).execute()
    return file_obj


def log_activity(message):
    """Log an activity message locally and to Drive."""
    try:
        os.makedirs("data", exist_ok=True)
        log_path = "data/daily_activity.log"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        with open(log_path, "a") as f:
            f.write(log_entry)
        
        # Sync the log to Drive (System Data folder)
        sync_data_file_to_drive(log_path)
    except Exception as e:
        print(f"Logging error: {e}")

def sync_data_file_to_drive(local_path):
    """Upload a specific local data file to the 'System Data' folder on Drive"""
    if not google_drive_is_ready():
        return False
    try:
        import os
        if not os.path.exists(local_path):
            return False
        file_name = os.path.basename(local_path)
        with open(local_path, "rb") as f:
            drive_file = upload_bytes_to_drive(["System Data"], file_name, f.read())
            if drive_file:
                # Store sync metadata
                if "last_sync" not in st.session_state:
                    st.session_state.last_sync = {}
                st.session_state.last_sync[local_path] = drive_file.get("id")
                return True
        return False
    except Exception as e:
        st.sidebar.error(f"Failed to sync {local_path} to Drive: {e}")
        return False


def pull_backend_data_from_drive():
    """Download all core data files from Drive to local data/ folder on startup.
    Includes backup logic and basic conflict avoidance."""
    if not google_drive_is_ready():
        log_activity("Sync check: Drive not ready.")
        return False
    
    service = get_google_drive_service()
    parent_id = ensure_drive_path(["System Data"])
    if not parent_id:
        st.sidebar.error("Could not find or create 'System Data' folder on Drive.")
        return False
        
    try:
        results = service.files().list(
            q=f"'{parent_id}' in parents and trashed = false",
            fields="files(id, name, modifiedTime)"
        ).execute()
        files = results.get("files", [])
        
        import io
        import shutil
        from datetime import datetime, timezone
        if not MediaIoBaseDownload:
            return False
            
        os.makedirs("data", exist_ok=True)
        
        success_count = 0
        for f in files:
            file_id = f['id']
            file_name = f['name']
            local_path = os.path.join("data", file_name)
            
            # Critical Conflict Prevention for Streamlit Cloud:
            # We trust Drive more on initial pull if the local file looks like a fresh clone.
            should_download = True
            if os.path.exists(local_path):
                local_mtime = datetime.fromtimestamp(os.path.getmtime(local_path), tz=timezone.utc)
                drive_mtime_str = f.get('modifiedTime', '').replace('Z', '+00:00')
                try:
                    drive_mtime = datetime.fromisoformat(drive_mtime_str)
                    
                    # Logic: If local is MUCH newer (more than 5 mins), assume it was saved in THIS session.
                    # Otherwise, IF Drive is newer OR local looks like original repo file, download.
                    now = datetime.now(timezone.utc)
                    age_seconds = (now - local_mtime).total_seconds()
                    
                    if age_seconds < 300: # Saved within last 5 minutes
                        if (local_mtime - drive_mtime).total_seconds() > 2:
                            should_download = False
                except:
                    pass

            if should_download:
                try:
                    request = service.files().get_media(fileId=file_id)
                    fh = io.BytesIO()
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while done is False:
                        status, done = downloader.next_chunk()
                    
                    # Store backup
                    if os.path.exists(local_path):
                        shutil.copy2(local_path, local_path + ".bak")

                    with open(local_path, "wb") as local_file:
                        local_file.write(fh.getvalue())
                    
                    # Set local mtime to match Drive
                    try:
                        drive_ts = datetime.fromisoformat(f.get('modifiedTime', '').replace('Z', '+00:00')).timestamp()
                        os.utime(local_path, (drive_ts, drive_ts))
                    except:
                        pass
                    
                    success_count += 1
                except Exception as e:
                    st.sidebar.warning(f"Failed to download {file_name}: {e}")
        
        log_activity(f"Pull completed: {success_count} files updated.")
        return True
    except Exception as e:
        st.sidebar.error(f"Drive pull error: {e}")
        if hasattr(get_google_drive_service, "clear"):
            get_google_drive_service.clear()
        return False


def save_uploaded_file(file_obj, local_path, drive_path_parts=None):
    import os

    os.makedirs(os.path.dirname(local_path), exist_ok=True)

    file_bytes = file_obj.getbuffer().tobytes()

    # Always save locally first
    with open(local_path, "wb") as f_out:
        f_out.write(file_bytes)

    # Mirror local directory structure for Google Drive upload
    local_dir = os.path.dirname(local_path)
    if local_dir:
        drive_path_parts = [part for part in os.path.normpath(local_dir).split(os.sep) if part]
    elif not drive_path_parts:
        drive_path_parts = ["DefaultProject", "General"]

    drive_result = None
    drive_error = None

    # Try uploading to Google Drive
    if google_drive_is_ready():
        try:
            drive_result = upload_bytes_to_drive(
                drive_path_parts,
                file_obj.name,
                file_bytes,
                getattr(file_obj, "type", None)
            )
        except Exception as exc:
            import traceback
            traceback.print_exc()
            drive_error = str(exc)
    else:
        debug_info = get_google_drive_debug_info()
        if not debug_info["root_id"]:
            drive_error = "Google Drive root folder ID is missing."
        elif not debug_info["credentials_ready"]:
            drive_error = "Google Drive credentials could not be loaded."
        else:
            drive_error = "Google Drive service not initialized."

    return drive_result, drive_error


LEGACY_PROJECT_NAMES = {
    "Truck Unloading Project": ["R&D Project", "Default Project"]
}

def get_project_storage_candidates(project_name):
    project_name = str(project_name or "").strip()
    if not project_name:
        return []
    return [project_name] + LEGACY_PROJECT_NAMES.get(project_name, [])

def get_existing_topic_dir(project_name, topic_name):
    for candidate_project in get_project_storage_candidates(project_name):
        topic_dir = os.path.join("pmo_storage", candidate_project, str(topic_name))
        if os.path.exists(topic_dir):
            return topic_dir
    return os.path.join("pmo_storage", str(project_name), str(topic_name))

def calculate_project_week(project_name, start_date, data_df):
    project_rows = data_df[data_df["Project"] == project_name].copy()
    if project_rows.empty:
        return 1

    project_start = pd.to_datetime(project_rows["Start Date"], errors="coerce").dropna()
    if project_start.empty:
        return 1

    start_dt = pd.to_datetime(start_date).tz_localize(None)
    master_start = project_start.min().tz_localize(None)
    delta_days = (start_dt - master_start).days
    return max(1, delta_days // 7 + 1)

def remember_week_context(project_name, week_num):
    st.session_state.preferred_weekly_project = project_name
    st.session_state.preferred_weekly_week = int(max(1, week_num))

def format_bullet_markdown(text):
    lines = []
    for raw_line in str(text or "").splitlines():
        cleaned = raw_line.strip()
        if not cleaned:
            continue
        if cleaned.startswith("-->"):
            cleaned = cleaned[3:].strip()
        elif cleaned.startswith("->"):
            cleaned = cleaned[2:].strip()
        cleaned = cleaned.lstrip("-*•").strip()
        if cleaned:
            lines.append(f"- {cleaned}")
    return "\n".join(lines) if lines else "-"

def format_single_line_text(text):
    parts = []
    for raw_line in str(text or "").splitlines():
        cleaned = raw_line.strip()
        if not cleaned:
            continue
        if cleaned.startswith("-->"):
            cleaned = cleaned[3:].strip()
        elif cleaned.startswith("->"):
            cleaned = cleaned[2:].strip()
        cleaned = cleaned.lstrip("-*•").strip()
        if cleaned:
            parts.append(cleaned)
    return " | ".join(parts)

def format_bullet_html(text):
    items = []
    for raw_line in str(text or "").splitlines():
        cleaned = raw_line.strip()
        if not cleaned:
            continue
        if cleaned.startswith("-->"):
            cleaned = cleaned[3:].strip()
        elif cleaned.startswith("->"):
            cleaned = cleaned[2:].strip()
        cleaned = cleaned.lstrip("-*•").strip()
        if cleaned:
            items.append(f"<li>{cleaned}</li>")
    return "<ul>" + "".join(items) + "</ul>" if items else "<ul><li>-</li></ul>"

def aggregate_topic_completion(topic_df):
    if topic_df.empty or "Completion %" not in topic_df.columns:
        return 0.0

    completion_values = pd.to_numeric(topic_df["Completion %"], errors="coerce").dropna()
    if completion_values.empty:
        return 0.0

    base_progress = float(completion_values.max())
    incremental_progress = float(completion_values[completion_values < base_progress].sum())
    return min(100.0, round(base_progress + incremental_progress, 1))

def order_topics(topic_values):
    topic_list = [str(topic).strip() for topic in topic_values if str(topic).strip() != ""]
    preferred_order = {topic: idx for idx, topic in enumerate(topics)}
    forced_last_topics = {"Container", "Objects"}
    return sorted(
        list(dict.fromkeys(topic_list)),
        key=lambda topic: (
            topic in forced_last_topics,
            preferred_order.get(topic, len(preferred_order)),
            topic.lower()
        )
    )

def build_topic_progress_df(task_df):
    if task_df.empty:
        return pd.DataFrame(columns=["Topic", "Completion %"])

    rows = []
    for topic, topic_df in task_df.groupby("Topic"):
        rows.append({
            "Topic": topic,
            "Completion %": aggregate_topic_completion(topic_df)
        })
    topic_progress_df = pd.DataFrame(rows)
    ordered_topics = order_topics(topic_progress_df["Topic"].tolist())
    topic_progress_df["Topic"] = pd.Categorical(topic_progress_df["Topic"], categories=ordered_topics, ordered=True)
    return topic_progress_df.sort_values("Topic").reset_index(drop=True)

def get_project_topics(project_name, data_df):
    if str(project_name).strip() == "":
        return []
    return order_topics(
        [
            str(topic)
            for topic in data_df[data_df["Project"] == project_name]["Topic"].dropna().unique()
        ]
    )

def get_milestone_topic(mil_info):
    explicit_topic = str(mil_info.get("topic", "")).strip()
    if explicit_topic:
        return explicit_topic

    task_topics = {
        str(task_info.get("topic", "")).strip()
        for task_info in mil_info.get("tasks", {}).values()
        if str(task_info.get("topic", "")).strip() != ""
    }
    if not task_topics:
        return ""
    if "All Topics" in task_topics or len(task_topics) > 1:
        return "All Topics"
    return next(iter(task_topics))

def get_milestone_progress(mil_info):
    pi = mil_info.get("progress_increase", 0)
    if isinstance(pi, dict):
        return sum(float(v or 0) for v in pi.values())
    return float(pi or 0)

def get_milestone_topic_increases(mil_info):
    """Return dict of {topic: increase} for per-topic progress. Handles both old and new format."""
    pi = mil_info.get("progress_increase", 0)
    if isinstance(pi, dict):
        return {k: float(v or 0) for k, v in pi.items()}
    # Legacy: single topic + single float
    topic = get_milestone_topic(mil_info)
    val = float(pi or 0)
    if topic and topic != "" and val > 0:
        return {topic: val}
    return {}

def get_planned_topic_adjustments(project_name, milestones):
    adjustments = {}
    if str(project_name).strip() == "":
        return adjustments

    for mil_info in milestones.values():
        milestone_project = str(mil_info.get("project_context", "")).strip()
        if milestone_project != project_name:
            continue
        if not bool(mil_info.get("completed", False)):
            continue
        topic_increases = get_milestone_topic_increases(mil_info)
        project_topics = get_project_topics(milestone_project, df)
        for topic, increase in topic_increases.items():
            if increase <= 0:
                continue
            if topic == "All Topics":
                for pt in project_topics:
                    adjustments[pt] = adjustments.get(pt, 0.0) + increase
            else:
                adjustments[topic] = adjustments.get(topic, 0.0) + increase

    return adjustments

def render_readonly_milestone(mil_id, mil_info):
    st.markdown(f"### 🎯 {mil_id}")
    topic_inc_display = get_milestone_topic_increases(mil_info)
    inc_lines = "  \n".join([f"**+{v:.0f}%** {t}" for t, v in topic_inc_display.items() if v > 0]) or "No increase set"
    info_col1, info_col2 = st.columns([4, 2])
    info_col1.markdown("**Description:**")
    info_col1.markdown(format_bullet_markdown(mil_info.get("description", "")))
    info_col2.markdown(
        f"**Project:** {mil_info.get('project_context', 'Not linked')}  \n"
        f"**Time:** {mil_info.get('time_needed', 0)} Hrs  \n"
        f"**Dates:** {mil_info.get('from_date', '')} to {mil_info.get('to_date', '')}  \n"
        f"**Topic Increases:**  \n{inc_lines}"
    )

    m_tasks = mil_info.get("tasks", {})
    with st.expander(f"Tasks {mil_id}", expanded=False):
        if not m_tasks:
            st.info("No tasks added yet for this milestone.")
        for t_id, t_info in m_tasks.items():
            head_col1, head_col2, head_col3 = st.columns([4.5, 1.4, 2.1])
            head_col1.markdown(f"**{t_id}**")
            head_col1.markdown(format_bullet_markdown(t_info.get("description", "")))
            head_col2.markdown(f"⏱️ {t_info.get('time_needed', 0)} Hrs")
            head_col3.markdown(
                f"📅 {t_info.get('from_date', '')} - {t_info.get('to_date','')}  \n"
                f"🏷️ {t_info.get('topic', 'Not linked')}"
            )

            if "errors" in t_info and len(t_info["errors"]) > 0:
                for idx, err in enumerate(t_info["errors"]):
                    err_cols = st.columns([8, 1])
                    timing_warn = " ⚠️ *(Timing May Vary)*" if err.get("timing_varied", False) else ""
                    err_cols[0].caption(f"⚠️ **Error/New Task {idx+1}:** {err['description']} *(+ {err['hours_spent']} hrs)*{timing_warn}")
                    solution_text = str(err.get("solution", "")).strip()
                    if solution_text:
                        err_cols[0].markdown("**Solution / Fix:**")
                        err_cols[0].markdown(format_bullet_markdown(solution_text))

            st.markdown("---")

def get_completed_milestone_total(project_name, milestones, topic_name="All Topics"):
    total = 0.0
    if str(project_name).strip() == "":
        return total

    for mil_info in milestones.values():
        if not bool(mil_info.get("completed", False)):
            continue
        if str(mil_info.get("project_context", "")).strip() != str(project_name).strip():
            continue
        milestone_topic = get_milestone_topic(mil_info)
        if topic_name != "All Topics" and milestone_topic not in [topic_name, "All Topics"]:
            continue
        total += get_milestone_progress(mil_info)

    return total




def render_completed_milestone(mil_id, mil_info, pm_data, data_df, project_options):
    render_readonly_milestone(mil_id, mil_info)

    if st.session_state.role != "Admin":
        return

    milestone_project = str(mil_info.get("project_context", "")).strip()
    milestone_topics = ["All Topics"] + get_project_topics(milestone_project, data_df) if milestone_project else ["All Topics"]

    action_col1, action_col2 = st.columns([1, 6])
    if action_col1.button("✏️ Edit Milestone", key=f"cm_edit_{mil_id}"):
        st.session_state[f"cm_edit_mode_{mil_id}"] = not st.session_state.get(f"cm_edit_mode_{mil_id}", False)
        st.rerun()
    if action_col2.button("🗑️ Delete Milestone", key=f"cm_delete_{mil_id}"):
        del pm_data[mil_id]
        save_planned_milestones(pm_data)
        st.rerun()

    if st.session_state.get(f"cm_edit_mode_{mil_id}", False):
        edit_col1, edit_col2, edit_col3, edit_col4, edit_col5 = st.columns([3, 1, 1, 2, 2])
        milestone_desc = edit_col1.text_area("Milestone Description", value=mil_info.get("description", ""), key=f"cm_desc_{mil_id}")
        edit_col2.text_input("Name", value=mil_id, disabled=True, key=f"cm_name_{mil_id}")
        milestone_time = edit_col3.number_input(
            "Time Needed (Hours)",
            min_value=0.0,
            step=0.5,
            value=float(mil_info.get("time_needed", 0)),
            key=f"cm_time_{mil_id}"
        )
        project_options_local = project_options.copy()
        if milestone_project and milestone_project not in project_options_local:
            project_options_local.append(milestone_project)
        project_index = project_options_local.index(milestone_project) if milestone_project in project_options_local else 0
        milestone_project_edit = edit_col5.selectbox("Project Context", project_options_local, index=project_index, key=f"cm_project_{mil_id}")

        date_col1, date_col2 = edit_col4.columns(2)
        milestone_start = date_col1.date_input(
            "From Date",
            value=pd.to_datetime(mil_info.get("from_date", datetime.now().date())).date(),
            key=f"cm_start_{mil_id}"
        )
        milestone_end = date_col2.date_input(
            "To Date",
            value=pd.to_datetime(mil_info.get("to_date", datetime.now().date())).date(),
            key=f"cm_end_{mil_id}"
        )

        topic_options = ["All Topics"] + get_project_topics(milestone_project_edit, data_df) if milestone_project_edit else ["All Topics"]
        current_milestone_topic = get_milestone_topic(mil_info) or "All Topics"
        if current_milestone_topic not in topic_options:
            topic_options.append(current_milestone_topic)
        impact_col1, impact_col2 = st.columns([2, 1])
        milestone_topic_edit = impact_col1.selectbox(
            "Milestone Topic",
            topic_options,
            index=topic_options.index(current_milestone_topic) if current_milestone_topic in topic_options else 0,
            key=f"cm_topic_{mil_id}"
        )
        milestone_pct = impact_col2.number_input(
            "Percentage Increase",
            min_value=0.0,
            max_value=100.0,
            value=get_milestone_progress(mil_info),
            step=1.0,
            key=f"cm_pct_{mil_id}"
        )

        save_col1, save_col2 = st.columns([1, 6])
        if save_col1.button("💾 Save Milestone", key=f"cm_save_{mil_id}"):
            if milestone_end < milestone_start:
                st.error("To Date cannot be earlier than From Date.")
            else:
                mil_info["description"] = str(milestone_desc)
                mil_info["time_needed"] = float(milestone_time)
                mil_info["from_date"] = milestone_start.strftime("%Y-%m-%d")
                mil_info["to_date"] = milestone_end.strftime("%Y-%m-%d")
                mil_info["project_context"] = milestone_project_edit
                mil_info["topic"] = milestone_topic_edit
                mil_info["progress_increase"] = float(milestone_pct)
                save_planned_milestones(pm_data)
                st.session_state[f"cm_edit_mode_{mil_id}"] = False
                st.success(f"Saved {mil_id}.")
                st.rerun()
        if save_col2.button("Cancel", key=f"cm_cancel_{mil_id}"):
            st.session_state[f"cm_edit_mode_{mil_id}"] = False
            st.rerun()

    st.markdown("**Admin Task Edit**")
    for t_id, t_info in mil_info.get("tasks", {}).items():
        with st.expander(f"✏️ Edit {t_id}", expanded=False):
            task_col1, task_col2, task_col3, task_col4 = st.columns([3, 1, 2, 2])
            edit_desc = task_col1.text_area("Task Description", value=t_info.get("description", ""), key=f"ct_desc_{mil_id}_{t_id}")
            edit_time = task_col2.number_input(
                "Time Needed",
                min_value=0.0,
                step=0.5,
                value=float(t_info.get("time_needed", 0)),
                key=f"ct_time_{mil_id}_{t_id}"
            )
            edit_start = task_col3.date_input(
                "From Date",
                value=pd.to_datetime(t_info.get("from_date", mil_info.get("from_date", datetime.now().date()))).date(),
                key=f"ct_start_{mil_id}_{t_id}"
            )
            edit_end = task_col4.date_input(
                "To Date",
                value=pd.to_datetime(t_info.get("to_date", mil_info.get("to_date", datetime.now().date()))).date(),
                key=f"ct_end_{mil_id}_{t_id}"
            )

            edit_topic = st.selectbox(
                "Topic",
                milestone_topics,
                index=milestone_topics.index(t_info.get("topic", "All Topics")) if t_info.get("topic", "All Topics") in milestone_topics else 0,
                key=f"ct_topic_{mil_id}_{t_id}"
            )

            save_col1, save_col2 = st.columns([1, 6])
            if save_col1.button("💾 Save Task", key=f"ct_save_{mil_id}_{t_id}"):
                if edit_end < edit_start:
                    st.error("To Date cannot be earlier than From Date.")
                else:
                    t_info["description"] = str(edit_desc)
                    t_info["time_needed"] = float(edit_time)
                    t_info["from_date"] = edit_start.strftime("%Y-%m-%d")
                    t_info["to_date"] = edit_end.strftime("%Y-%m-%d")
                    t_info["project"] = milestone_project
                    t_info["topic"] = edit_topic
                    save_planned_milestones(pm_data)
                    st.success(f"Saved {t_id}.")
                    st.rerun()
            if save_col2.button("🗑️ Delete Task", key=f"ct_delete_{mil_id}_{t_id}"):
                del mil_info["tasks"][t_id]
                save_planned_milestones(pm_data)
                st.rerun()


# --- INITIAL SYNC ---
if "data_synced" not in st.session_state:
    with st.spinner("Syncing data from Drive..."):
        pull_backend_data_from_drive()
    st.session_state.data_synced = True

# --- DIAGNOSTICS AT TOP OF PAGE ---
if not google_drive_is_ready():
    st.sidebar.error("⚠️ Drive Disconnected")
    with st.expander("🛠️ Google Drive Debug Info", expanded=True):
        st.error("Google Drive is not connected. This is why 'Refresh' is failing.")
        creds = _load_google_drive_credentials()
        if creds is None:
            st.warning("Reason: No valid credentials found. Please check your Streamlit Secrets.")
            st.info("Ensure you have [service_account] block in your Secrets.")
        else:
            st.success("Credentials loaded, but service initialization failed.")
else:
    st.sidebar.success("✅ Drive Connected")

with st.sidebar:
    st.markdown("---")
    with st.expander("🔄 Data Sync Settings"):
        if st.button("🔄 Refresh from Cloud"):
            with st.spinner("Refreshing..."):
                try:
                    ok = pull_backend_data_from_drive()
                    if ok:
                        st.success("Successfully refreshed from Cloud.")
                        st.rerun()
                    else:
                        st.error("Refresh failed. Check the sidebar for details.")
                except Exception as e:
                    st.error(f"Sync Button Error: {e}")
        
        # Display last updated timestamp for local storage
        if os.path.exists("data/tasks.xlsx"):
            mtime = os.path.getmtime("data/tasks.xlsx")
            st.caption(f"Local storage last updated: {datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')}")
    st.markdown("---")

df = load_data()

# --- DYNAMIC PROJECTS & TOPICS ---
# Isolate Projects
base_projects = ["Truck Unloading Project"]
saved_projects = [str(p) for p in df['Project'].dropna().unique() if str(p).strip() != ""]
projects = list(dict.fromkeys(base_projects + saved_projects))

# Isolate Employees
base_emps = ["Unassigned", "Employee 1", "Employee 2", "Employee 3", "Employee 4"]
saved_emps = [str(e) for e in df.get('Employee', pd.Series([])).dropna().unique() if str(e).strip() != ""]
employees = list(dict.fromkeys(base_emps + saved_emps))

# Isolate Topics
base_topics = ["Robot", "Vision System", "Conveyor", "AGV", "EOAT", "Vacuum System", "Container", "Objects"]
saved_topics = [str(t) for t in df['Topic'].dropna().unique() if str(t).strip() != ""]
topics = list(dict.fromkeys(base_topics + saved_topics))

STATUS_OPTIONS = ["Planned", "In Progress", "Completed", "Delayed"]

# --- INITIALIZE IMAGES IN SESSION ---
if 'topic_images' not in st.session_state:
    st.session_state.topic_images = {}  
if "last_drive_status" not in st.session_state:
    st.session_state.last_drive_status = None

# --- INITIALIZE AUTHENTICATION ---
if 'role' not in st.session_state:
    st.session_state.role = None
if 'auth_name' not in st.session_state:
    st.session_state.auth_name = None
if 'preferred_weekly_project' not in st.session_state:
    st.session_state.preferred_weekly_project = None
if 'preferred_weekly_week' not in st.session_state:
    st.session_state.preferred_weekly_week = None

if st.session_state.role is None:
    st.title("🔒 PMO Authentication Gateway")
    st.markdown("Please identify your security clearance to access the workspace.")
    
    st.divider()
    auth_col1, auth_col2 = st.columns(2)
    with auth_col1:
        st.subheader("👨‍💻 Employee Verification")
        emp_name = st.text_input("Enter your full name", key="emp_n")
        if st.button("Enter PMO (Employee)", use_container_width=True):
            if str(emp_name).strip() != "":
                st.session_state.role = "Employee"
                st.session_state.auth_name = str(emp_name).strip()
                st.rerun()
            else:
                st.error("Identification required.")
                
    with auth_col2:
        st.subheader("🛡️ Administrator Override")
        admin_pass = st.text_input("Enter Master Node Key", type="password", key="adm_p")
        if st.button("Enter PMO (Admin)", use_container_width=True):
            if admin_pass == "effica123":
                st.session_state.role = "Admin"
                st.session_state.auth_name = "Administrator"
                st.rerun()
            else:
                st.error("Invalid security clearance.")
    st.stop()

# --- SIDEBAR ---
st.sidebar.title("BASELINE PMO")
st.sidebar.caption(f"Logged in as: **{st.session_state.auth_name}** ({st.session_state.role})")
if st.sidebar.button("🚪 Logout Data Session"):
    st.session_state.role = None
    st.rerun()
st.sidebar.divider()
nav_options = ["Dashboard", "Weekly Performance", "Tasks & Milestones", "Planned Milestones", "Image Gallery", "Competitors & Research"]
if st.session_state.role == "Admin":
    nav_options.append("Document Drive")
page = st.sidebar.selectbox("Navigation", nav_options)

# ─────────────────────────────────────────────
# DASHBOARD PAGE
# ─────────────────────────────────────────────
if page == "Dashboard":
    st.title("R&D Project Overview & Analytics")
    pm_data = load_planned_milestones()
    
    # Select Project Context
    selected_project = st.selectbox("🌐 Select View Context (Project Filter)", projects)
    st.divider()

    # Filter dataframe exclusively to this project
    proj_df = df[df["Project"] == selected_project]
    
    # Topics specific to this project
    proj_topics = order_topics(proj_df['Topic'].dropna().unique())
    if not proj_topics:
        proj_topics = topics[:4]  # Default placeholders if empty
    topic_adjustments = get_planned_topic_adjustments(selected_project, pm_data)

    st.subheader(f"Dashboard » {selected_project}")

    def render_topic_files(t_proj, t_topic, btn_key=""):
        with st.popover("📂 Topic Files", use_container_width=True):
            topic_dir = get_existing_topic_dir(t_proj, t_topic)
            meta_path = os.path.join(topic_dir, ".metadata.json")
            
            topic_meta = {}
            if os.path.exists(meta_path):
                try:
                    with open(meta_path, "r") as mf:
                        topic_meta = json.load(mf)
                except:
                    topic_meta = {}

            # 1. Upload Section at the TOP
            st.markdown("##### Upload Document / Link")
            up_type = st.radio("Type", ["File", "Link"], horizontal=True, key=f"uptype_{btn_key}_{t_topic}")
            
            cnt_key = f"cnt_{btn_key}_{t_topic}"
            if cnt_key not in st.session_state:
                st.session_state[cnt_key] = 0
            cnt = st.session_state[cnt_key]
            
            up_note_key = f"upnote_{btn_key}_{t_topic}_{cnt}"
            uplink_key = f"uplink_{btn_key}_{t_topic}_{cnt}"
            upfile_key = f"upfile_{btn_key}_{t_topic}_{cnt}"
            
            up_note = st.text_input("Small Note (Optional)", key=up_note_key)
            
            if up_type == "File":
                up_file = st.file_uploader("Upload File", key=upfile_key)
                if st.button("Upload File", key=f"btn_up_{btn_key}_{t_topic}"):
                    if up_file:
                        os.makedirs(topic_dir, exist_ok=True)
                        save_path = os.path.join(topic_dir, up_file.name)
                        drive_file, drive_error = save_uploaded_file(
                            up_file,
                            save_path,
                            [t_proj, t_topic]
                        )
                            
                        if "files" not in topic_meta:
                            topic_meta["files"] = {}
                        topic_meta["files"][up_file.name] = {
                            "note": up_note,
                            "drive_url": drive_file.get("webViewLink", "") if drive_file else "",
                            "drive_file_id": drive_file.get("id", "") if drive_file else ""
                        }
                        with open(meta_path, "w") as mf:
                            json.dump(topic_meta, mf, indent=4)
                        if drive_error:
                            print(f"DRIVE UPLOAD ERROR: {drive_error}")
                            st.error(f"Saved locally, but Google Drive sync failed: {drive_error}")
                        else:
                            st.success("File uploaded to Google Drive successfully!")
                            import time
                            time.sleep(1)
                            st.session_state[cnt_key] += 1
                            st.rerun()
                    else:
                        st.error("Please select a file to upload.")
            else:
                up_link = st.text_input("Enter URL", key=uplink_key)
                if st.button("Add Link", key=f"btn_link_{btn_key}_{t_topic}"):
                    if up_link:
                        os.makedirs(topic_dir, exist_ok=True)
                        if "links" not in topic_meta:
                            topic_meta["links"] = {}
                        import uuid
                        link_id = str(uuid.uuid4())
                        topic_meta["links"][link_id] = {"url": up_link, "note": up_note}
                        with open(meta_path, "w") as mf:
                            json.dump(topic_meta, mf, indent=4)
                            
                        st.session_state[cnt_key] += 1
                        st.rerun()
                    else:
                        st.error("Please provide a valid URL.")

            st.divider()

            # 2. Display Section BELOW
            if up_type == "File":
                with st.expander("📂 Show Attached Files", expanded=False):
                    has_files = False
                    
                    # Combine physical files + metadata files
                    files_to_show = []
                    if os.path.exists(topic_dir):
                        files_to_show = [f for f in os.listdir(topic_dir) if f != ".metadata.json"]
                    
                    # Also include files from metadata that aren't downloaded physically
                    meta_files = topic_meta.get("files", {})
                    for f_name in meta_files.keys():
                        if f_name not in files_to_show:
                            files_to_show.append(f_name)

                    if files_to_show:
                        has_files = True
                        for file_item in files_to_show:
                            file_path = os.path.join(topic_dir, file_item)
                            f_info = topic_meta.get("files", {}).get(file_item, {})
                            f_note = f_info.get("note", "")
                            d_url = f_info.get("drive_url", "")
                            
                            f_col1, f_col2 = st.columns([4, 1])
                            with f_col1:
                                if os.path.exists(file_path):
                                    if file_item.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                                        st.image(file_path, caption=f"{file_item}" + (f" - {f_note}" if f_note else ""), use_container_width=True)
                                    else:
                                        with open(file_path, "rb") as bf:
                                            lbl_txt = f"📄 {file_item}" + (f" - {f_note}" if f_note else "")
                                            st.download_button(label=lbl_txt, data=bf, file_name=file_item, mime="application/octet-stream", key=f"dl_{btn_key}_{t_topic}_{file_item}")
                                else:
                                    if d_url:
                                        lbl_txt = f"☁️ [View missing local file on Drive: {file_item}]({d_url})" + (f" - {f_note}" if f_note else "")
                                        st.markdown(lbl_txt)
                                    else:
                                        st.warning(f"⚠️ {file_item} (Missing locally without Drive link)")

                            with f_col2:
                                if st.button("🗑️", key=f"del_{btn_key}_{t_topic}_{file_item}", help="Delete File"):
                                    if os.path.exists(file_path):
                                        try:
                                            os.remove(file_path)
                                        except:
                                            pass
                                    if "files" in topic_meta and file_item in topic_meta["files"]:
                                        del topic_meta["files"][file_item]
                                        with open(meta_path, "w") as mf:
                                            json.dump(topic_meta, mf, indent=4)
                                    st.rerun()

                    if not has_files:
                        st.info(f"No files attached to '{t_topic}' yet.")

            else:
                with st.expander("🔗 Show Attached Links", expanded=False):
                    has_links = False
                    links_to_show = topic_meta.get("links", {})
                    if links_to_show:
                        has_links = True
                        for link_id, link_info in list(links_to_show.items()):
                            l_url = link_info.get("url", "")
                            l_note = link_info.get("note", "")
                            
                            l_col1, l_col2 = st.columns([4, 1])
                            with l_col1:
                                lbl_txt = f"🔗 [{l_url}]({l_url})" + (f" - *{l_note}*" if l_note else "")
                                st.markdown(lbl_txt)
                            with l_col2:
                                if st.button("🗑️", key=f"del_link_{btn_key}_{t_topic}_{link_id}", help="Delete Link"):
                                    del topic_meta["links"][link_id]
                                    with open(meta_path, "w") as mf:
                                        json.dump(topic_meta, mf, indent=4)
                                    st.rerun()
                    if not has_links:
                        st.info(f"No links attached to '{t_topic}' yet.")

    # 1. TOP GAUGES + optional subsystem images (DYNAMICALLY SIZED BY TOPICS IN PROJECT)
    cols = st.columns(max(len(proj_topics), 1))
    for i, topic in enumerate(proj_topics):
        topic_tasks = proj_df[
            (proj_df["Topic"] == topic) &
            (proj_df["Hidden"] == False)
        ]
        avg_comp = min(100.0, aggregate_topic_completion(topic_tasks) + topic_adjustments.get(topic, 0.0))

        with cols[i]:
            img_path = f"data/topic_images/{selected_project}_{topic}.png"
            if os.path.exists(img_path):
                # Fallback to local disk image tracking
                st.image(img_path, use_container_width=True)
            elif topic in st.session_state.topic_images:
                st.image(st.session_state.topic_images[topic], use_container_width=True)

            fig = go.Figure(go.Indicator(
                mode="gauge+number",
                value=avg_comp,
                title={'text': f"<b>{topic}</b>", 'font': {'size': 12}},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#2ecc71"}}
            ))
            fig.update_layout(height=180, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)
            
            # Attach Document Drive Popover below Gauge
            render_topic_files(selected_project, topic, btn_key="gauge")

    st.divider()

    # 2. TASK PROGRESSION CHART
    st.subheader(f"Topic Progression » {selected_project}")
    visible_tasks = proj_df[proj_df["Hidden"] == False]
    if not visible_tasks.empty:
        topic_avg = build_topic_progress_df(visible_tasks)
        if not topic_avg.empty:
            topic_avg["Completion %"] = topic_avg.apply(
                lambda row: min(100.0, float(row["Completion %"]) + topic_adjustments.get(str(row["Topic"]), 0.0)),
                axis=1
            )
        
        fig_bar = px.bar(
            topic_avg,
            x="Completion %", y="Topic",
            color="Topic", orientation='h',
            text="Completion %",
            category_orders={"Topic": proj_topics}
        )
        fig_bar.update_xaxes(range=[0, 100], showticklabels=False, title="")
        fig_bar.update_yaxes(
            autorange="reversed",
            categoryorder="array",
            categoryarray=proj_topics
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info(f"No tasks added yet for '{selected_project}'. Go to 'Tasks & Milestones' to feed data.")

    st.divider()

    # 3. BOTTOM DETAIL GRID
    st.divider()
    
    col_hdr, col_tgl = st.columns([4, 1])
    col_hdr.subheader("Project Context & Analytics")
    edit_notes = col_tgl.toggle("📝 Enable Grid Editor", value=False)
    
    notes_db = load_notes()
    if selected_project not in notes_db:
        notes_db[selected_project] = {
            "Topics": {},
            "Project_Issues": "**Critical:** AGV Mapping error in Zone B\n**Low:** Vacuum suction cup wear",
            "Project_Plans": "- 2026 Q3: Full Site Expansion\n- Multi-fleet cloud sync"
        }
    proj_notes = notes_db[selected_project]

    col_l, col_r = st.columns([3, 1])
    with col_l:
        st.markdown("**Topic Context Library**")
        
        # Automatically sync topics drawn from actual project tasks
        for t in proj_topics:
            if t not in proj_notes["Topics"]:
                proj_notes["Topics"][t] = {
                    "Major": "",
                    "Problematic": "",
                    "Future": ""
                }
        active_topics = proj_topics
        
        if not active_topics:
            st.info("No topic context cards exist yet. Enable Grid Editor above to launch one!")
            
        # Draw dynamic grid iterating strictly over active configuration
        for i in range(0, len(active_topics), 3):
            grid_cols = st.columns(3)
            for j, topic in enumerate(active_topics[i:i+3]):
                with grid_cols[j]:
                    tn = proj_notes["Topics"][topic]
                    
                    if edit_notes:
                        st.markdown(f"#### 📦 {topic}")
                        new_maj = st.text_area("Completed task", tn.get("Major", ""), key=f"nm_{topic}", height=70)
                        new_prob = st.text_area("In progress", tn.get("Problematic", ""), key=f"np_{topic}", height=70)
                        new_fut = st.text_area("Future Phase", tn.get("Future", ""), key=f"nf_{topic}", height=70)
                        
                        tn["Major"] = new_maj
                        tn["Problematic"] = new_prob
                        tn["Future"] = new_fut
                        st.divider()
                    else:
                        maj_html = format_bullet_html(tn.get("Major", ""))
                        prob_html = format_bullet_html(tn.get("Problematic", ""))
                        fut_html = format_bullet_html(tn.get("Future", ""))
                        st.markdown(f"""
                        <div class="card">
                            <h4>📦 {topic}</h4>
                            <b>Completed task:</b>{maj_html}
                            <b>In progress:</b>{prob_html}
                            <b>Future Phase:</b>{fut_html}
                        </div>
                        """, unsafe_allow_html=True)

                        render_topic_files(selected_project, topic, btn_key="grid")

    with col_r:
        if edit_notes:
            st.markdown("#### ⚠️ Issues & Plans")
            new_iss = st.text_area("Project Issues", proj_notes["Project_Issues"], height=100)
            new_pl = st.text_area("Further Plans", proj_notes["Project_Plans"], height=100)
            proj_notes["Project_Issues"] = new_iss
            proj_notes["Project_Plans"] = new_pl
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 Save All Notes to System", use_container_width=True):
                save_notes(notes_db)
                st.success("Updates locked globally!")
        else:
            st.markdown("#### ⚠️ Project Issues")
            st.markdown(format_bullet_markdown(proj_notes["Project_Issues"]))
            st.markdown("#### 🚀 Further Plans")
            st.markdown(format_bullet_markdown(proj_notes["Project_Plans"]))


# ─────────────────────────────────────────────
# WEEKLY PERFORMANCE PAGE
# ─────────────────────────────────────────────
elif page == "Weekly Performance":
    st.header("Weekly Performance Analytics")
    
    sel_col1, sel_col2 = st.columns(2)
    preferred_project = st.session_state.get("preferred_weekly_project")
    project_index = projects.index(preferred_project) if preferred_project in projects else 0
    selected_project = sel_col1.selectbox("🌐 Select View Context (Project Filter)", projects, index=project_index)
    
    # Get topics available in this project
    proj_topics_list = order_topics(df[df["Project"] == selected_project]['Topic'].dropna().unique())
    selected_topic = sel_col2.selectbox("🏷️ Select Topic Filter", ["All Topics"] + proj_topics_list)
    
    st.divider()
    
    proj_df = df[df["Project"] == selected_project]
    if selected_topic != "All Topics":
        proj_df = proj_df[proj_df["Topic"] == selected_topic]
    
    st.subheader(f"Week-by-Week Comparison » {selected_project}" + (f" » {selected_topic}" if selected_topic != "All Topics" else ""))
    
    # Identify available weeks in the selected project
    proj_weeks = proj_df["Week"].dropna().unique() if "Week" in proj_df.columns else [1]
    proj_weeks = sorted([int(w) for w in proj_weeks])
    if not proj_weeks:
        proj_weeks = [1]
        
    # Two dropdowns for specific week comparison (NOT a range — only the two selected weeks are shown)
    preferred_week = st.session_state.get("preferred_weekly_week")
    default_start_week = preferred_week if preferred_week in proj_weeks else proj_weeks[0]

    st.caption("ℹ️ Select **two specific weeks** to compare side-by-side. Only those exact weeks are shown (not the range between them).")
    wcol1, wcol2 = st.columns(2)
    start_wk = wcol1.selectbox("📊 Week 1 (Select)", options=proj_weeks, index=proj_weeks.index(default_start_week))

    default_end_week = preferred_week if preferred_week in proj_weeks else proj_weeks[-1]
    end_wk = wcol2.selectbox("📊 Week 2 (Compare)", options=proj_weeks, index=proj_weeks.index(default_end_week))

    if "Week" in proj_df.columns:
        week_df_start = proj_df[(proj_df["Week"] == start_wk) & (proj_df["Hidden"] == False)]
        week_df_end = proj_df[(proj_df["Week"] == end_wk) & (proj_df["Hidden"] == False)]
    else:
        week_df_start = pd.DataFrame()
        week_df_end = pd.DataFrame()

    if not week_df_start.empty or not week_df_end.empty:
        if selected_topic == "All Topics":
            st.markdown("**Topic Progress per Week**")
            frames = []
            if not week_df_start.empty:
                avg_s = week_df_start.groupby("Topic")["Completion %"].mean().reset_index()
                avg_s["Week_Label"] = f"Wk {start_wk}"
                frames.append(avg_s)
            if not week_df_end.empty and end_wk != start_wk:
                avg_e = week_df_end.groupby("Topic")["Completion %"].mean().reset_index()
                avg_e["Week_Label"] = f"Wk {end_wk}"
                frames.append(avg_e)
            if frames:
                combined = pd.concat(frames, ignore_index=True)
                combined["Topic"] = pd.Categorical(combined["Topic"], categories=proj_topics_list, ordered=True)
                combined = combined.sort_values(["Topic", "Week_Label"]).reset_index(drop=True)
                fig_week = px.bar(combined, x="Week_Label", y="Completion %", color="Topic", barmode="group",
                                  text_auto='.1f', labels={'Completion %': 'Average Completion (%)', 'Week_Label': 'Project Week'},
                                  category_orders={"Topic": proj_topics_list, "Week_Label": [f"Wk {start_wk}", f"Wk {end_wk}"] if end_wk != start_wk else [f"Wk {start_wk}"]})
                fig_week.update_layout(yaxis=dict(range=[0, 100]), height=400)
                st.plotly_chart(fig_week, use_container_width=True)
        else:
            st.markdown(f"**{selected_topic} Progress per Week**")
            frames = []
            if not week_df_start.empty:
                avg_s = week_df_start.groupby("Week")["Completion %"].mean().reset_index()
                avg_s["Week_Label"] = f"Wk {start_wk}"
                frames.append(avg_s)
            if not week_df_end.empty and end_wk != start_wk:
                avg_e = week_df_end.groupby("Week")["Completion %"].mean().reset_index()
                avg_e["Week_Label"] = f"Wk {end_wk}"
                frames.append(avg_e)
            if frames:
                combined = pd.concat(frames, ignore_index=True)
                fig_week = px.bar(combined, x="Week_Label", y="Completion %", text_auto='.1f',
                                  labels={'Completion %': 'Average Completion (%)', 'Week_Label': 'Project Week'},
                                  color="Week_Label", color_discrete_sequence=px.colors.qualitative.Vivid)
                fig_week.update_layout(yaxis=dict(range=[0, 100]), height=400, showlegend=False)
                st.plotly_chart(fig_week, use_container_width=True)
    else:
        st.info("No active tasks found for the selected weeks.")

# ─────────────────────────────────────────────
# TASKS & MILESTONES PAGE
# ─────────────────────────────────────────────
elif page == "Tasks & Milestones":
    st.header("Task Management")
    pm_data = load_planned_milestones()

    # --- ADD NEW TASK FORM ---
    with st.expander("➕ Add New Task", expanded=False):
        st.markdown("**1. Project Assignment**")
        pc1, pc2 = st.columns(2)
        t_proj_sel = pc1.selectbox("Map to Existing Project", projects, key="n_proj_sel")
        t_proj_new = pc2.text_input("...OR Create entirely New Project", help="Will generate a pristine new dashboard context view", key="n_proj_new")
        
        actual_proj = str(t_proj_new).strip() if str(t_proj_new).strip() != "" else str(t_proj_sel).strip()
        
        # Calculate dynamic elements based on selected project
        proj_tasks = df[df["Project"] == actual_proj]
        proj_start_master_dt = None
        if not proj_tasks.empty:
            proj_start_master_dt = pd.to_datetime(proj_tasks["Start Date"]).min()
            proj_end_master_dt = pd.to_datetime(proj_tasks["End Date"]).max()
            total_weeks = max(1, (proj_end_master_dt - proj_start_master_dt).days // 7 + 1)
            st.info(f"📅 **Project Timeline**: {total_weeks} Expected Weeks")
        else:
            st.info(f"📅 **Project Timeline**: Brand New Project - Initializing")

        st.markdown("**2. Sub-Topic Classification**")
        tc1, tc2 = st.columns(2)
        t_topic_new = tc1.text_input("Topic Name", help="Enter topic name to map this task to", key="n_topic_new")
        t_topic_sel = tc2.selectbox("...OR Map to Existing Topic", [""] + topics, key="n_topic_sel")
            
        st.markdown("**3. Task Constraints / Assignment**")
        ac1, ac2 = st.columns(2)
        t_name   = ac1.text_area("Task Specification (Description)", key="n_name", height=120)
        
        t_emp_sel = ac2.selectbox("Employee Assigned", employees, key="n_emp_sel")
        t_emp_new = ac2.text_input("...OR Add New Employee Name", help="Will override dropdown above", key="n_emp_new")
        
        sc1, sc2, sc3, sc4, sc5 = st.columns(5)
        t_start  = sc1.date_input("Start Date", datetime.now(), key="n_start")
        t_end    = sc2.date_input("End Date", datetime.now() + timedelta(days=7), key="n_end")
        
        # Dynamic Derived Week
        derived_week = calculate_project_week(actual_proj, t_start, df)

        if st.session_state.role == "Admin":
            t_week = sc3.number_input("Project Week", min_value=1, max_value=2000, value=derived_week, step=1, key="n_week", help="Defaults from Project Start Date, but Admin can adjust it.")
        else:
            t_week = sc3.number_input("Project Week", value=derived_week, disabled=True, help="Automatically calculates offset from Project Start Date")
        t_comp = sc4.number_input("% Done", min_value=0, max_value=100, value=0, step=1, key="n_comp")
        t_status = sc5.selectbox("Status", STATUS_OPTIONS, key="n_status")
        st.caption(f"Selected dates map to Project Week {derived_week}.")

        st.markdown("**4. Milestone Strategy & Planning**")
        m_col1, m_col2 = st.columns([2, 1])
        t_milestone = m_col1.text_area("Milestone Strategy (How to accomplish this task)", key="n_milestone", height=100)
        t_files = m_col2.file_uploader("Attach Milestone Documents", accept_multiple_files=True, key="n_files")

        if st.button("💾 Submit & Save Everything to Excel"):
            actual_topic = str(t_topic_new).strip() if str(t_topic_new).strip() != "" else str(t_topic_sel).strip()
            actual_emp = str(t_emp_new).strip() if str(t_emp_new).strip() != "" else str(t_emp_sel).strip()
            
            if actual_proj == "" or actual_topic == "" or str(t_name).strip() == "":
                st.error("Please fill in the Project, Topic, and Task Name to lock the structure.")
            elif t_end < t_start:
                st.error("End Date cannot be earlier than Start Date.")
            else:
                new_task = pd.DataFrame([{
                    "Project": actual_proj, 
                    "Topic": actual_topic, 
                    "Task Name": t_name,
                    "Start Date": t_start.strftime("%Y-%m-%d"), 
                    "End Date": t_end.strftime("%Y-%m-%d"),
                    "Completion %": t_comp, 
                    "Status": t_status,
                    "Employee": actual_emp,
                    "Week": int(t_week),
                    "Hidden": False,
                    "Milestone_Text": t_milestone,
                    "Milestone_Author_Name": st.session_state.auth_name,
                    "Milestone_Role": st.session_state.role
                }])
                # Append strictly to DataFrame
                df = pd.concat([df, new_task], ignore_index=True)
                remember_week_context(actual_proj, int(t_week))
                
                # Handle File Uploads specifically for this milestone
                if t_files:
                    m_dir = f"pmo_storage/{actual_proj}/{actual_topic}/{t_name}_Milestone"
                    os.makedirs(m_dir, exist_ok=True)
                    drive_failures = []
                    for f in t_files:
                        _, drive_error = save_uploaded_file(
                            f,
                            os.path.join(m_dir, f.name),
                            [actual_proj, actual_topic, f"{t_name}_Milestone"]
                        )
                        if drive_error:
                            drive_failures.append(f"{f.name}: {drive_error}")
                    if drive_failures:
                        st.warning("Milestone files saved locally, but some Drive syncs failed: " + " | ".join(drive_failures[:3]))
                            
                save_data(df)
                st.success(f"Successfully integrated '{t_name}' under '{actual_topic}' in '{actual_proj}'!")
                st.rerun()

    st.divider()

    completed_milestones = {
        mil_id: mil_info
        for mil_id, mil_info in pm_data.items()
        if bool(mil_info.get("completed", False))
    }
    if completed_milestones:
        st.subheader("Completed Milestones")
        completed_projects = sorted(
            {
                str(mil_info.get("project_context", "")).strip()
                for mil_info in completed_milestones.values()
                if str(mil_info.get("project_context", "")).strip() != ""
            }
        )
        cm_col1, cm_col2, cm_col3 = st.columns([2, 2, 2])
        selected_completed_project = cm_col1.selectbox(
            "Completed Milestone Project",
            ["Select Project"] + completed_projects,
            key="cm_filter_project"
        )
        if selected_completed_project == "Select Project":
            completed_topic_options = ["All Topics"]
        else:
            completed_topic_set = set(get_project_topics(selected_completed_project, df))
            for mil_info in completed_milestones.values():
                if str(mil_info.get("project_context", "")).strip() != selected_completed_project:
                    continue
                milestone_topic = get_milestone_topic(mil_info)
                if milestone_topic:
                    completed_topic_set.add(milestone_topic)
            completed_topic_options = ["All Topics"] + sorted(completed_topic_set)
        selected_completed_topic = cm_col2.selectbox(
            "Completed Milestone Topic",
            completed_topic_options,
            key="cm_filter_topic"
        )
        added_total = 0.0 if selected_completed_project == "Select Project" else get_completed_milestone_total(
            selected_completed_project,
            completed_milestones,
            selected_completed_topic
        )
        cm_col3.metric("Added Completion %", f"{added_total:.1f}%")

        if selected_completed_project == "Select Project":
            st.info("Select a project to view completed milestones.")
        else:
            filtered_completed_milestones = []
            for mil_id, mil_info in completed_milestones.items():
                if str(mil_info.get("project_context", "")).strip() != selected_completed_project:
                    continue
                if selected_completed_topic != "All Topics":
                    milestone_topic = get_milestone_topic(mil_info)
                    if selected_completed_topic not in [milestone_topic, "All Topics"]:
                        continue
                filtered_completed_milestones.append((mil_id, mil_info))

            if not filtered_completed_milestones:
                st.info("No completed milestones match this project/topic filter.")
            else:
                for mil_id, mil_info in filtered_completed_milestones:
                    with st.container():
                        render_completed_milestone(mil_id, mil_info, pm_data, df, projects)
                        st.divider()

    # --- FILTER SECTION ---
    st.subheader("📋 Task List")
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([2, 2, 2, 2])

    with filter_col1:
        project_filter_options = ["Select Project"] + projects
        filter_project = st.selectbox("Isolate by Project", project_filter_options)
    with filter_col2:
        if filter_project == "Select Project":
            topic_filter_options = ["Select Topic"]
        else:
            project_topics = sorted(
                [
                    str(t)
                    for t in df[df["Project"] == filter_project]["Topic"].dropna().unique()
                    if str(t).strip() != ""
                ]
            )
            topic_filter_options = ["All Topics"] + project_topics
        filter_topic = st.selectbox("Filter by System/Topic", topic_filter_options)
    with filter_col3:
        filter_status = st.selectbox("Filter by Status", ["All"] + STATUS_OPTIONS)
    with filter_col4:
        show_hidden = st.checkbox("Show Hidden Tasks", value=False)

    if filter_project == "Select Project" or filter_topic == "Select Topic":
        st.info("Select a project to view the task list. You can then keep `All Topics` or choose a specific topic.")
    else:
        # Apply filters
        display_df = df.copy()
        if not show_hidden:
            display_df = display_df[display_df["Hidden"] == False]
        display_df = display_df[display_df["Project"] == filter_project]
        if filter_topic != "All Topics":
            display_df = display_df[display_df["Topic"] == filter_topic]
        if filter_status != "All":
            display_df = display_df[display_df["Status"] == filter_status]

        filtered_indices = display_df.index.tolist()

        if display_df.empty:
            st.info("No tasks align with current filter contexts.")
        else:
        # --- TASK TABLE WITH EDIT / DELETE / HIDE ---
        # Admin sees Edit / Hide / Del columns; Employee sees read-only view
            if st.session_state.role == "Admin":
                hcols = st.columns([0.6, 1.5, 1.3, 2.5, 1.2, 1.2, 1.2, 0.6, 0.8, 1.2, 0.6, 0.6, 0.6])
                headers = ["#", "Project", "Topic", "Task Name", "Start", "End", "Employee", "Wk", "Done %", "Status", "Edit", "Hide", "Del"]
            else:
                hcols = st.columns([0.6, 1.5, 1.3, 2.5, 1.2, 1.2, 1.2, 0.6, 0.8, 1.2])
                headers = ["#", "Project", "Topic", "Task Name", "Start", "End", "Employee", "Wk", "Done %", "Status"]
            for h, col in zip(headers, hcols):
                col.markdown(f"**{h}**")

            st.markdown("---")

            for row_num, idx in enumerate(filtered_indices):
                row = df.loc[idx]
                is_hidden = bool(row["Hidden"])
                task_name_toggle_key = f"show_task_name_{idx}"
                if task_name_toggle_key not in st.session_state:
                    st.session_state[task_name_toggle_key] = False

                if st.session_state.role == "Admin":
                    rcols = st.columns([0.6, 1.5, 1.3, 2.5, 1.2, 1.2, 1.2, 0.6, 0.8, 1.2, 0.6, 0.6, 0.6])
                else:
                    rcols = st.columns([0.6, 1.5, 1.3, 2.5, 1.2, 1.2, 1.2, 0.6, 0.8, 1.2])
                rcols[0].write(row_num + 1)
                rcols[1].write(row["Project"])
                rcols[2].write(row["Topic"])
                if rcols[3].button(
                    "Hide Task Name" if st.session_state[task_name_toggle_key] else "View Task Name",
                    key=f"btn_task_name_{idx}"
                ):
                    st.session_state[task_name_toggle_key] = not st.session_state[task_name_toggle_key]
                    st.rerun()
                rcols[4].write(str(row["Start Date"]))
                rcols[5].write(str(row["End Date"]))
                rcols[6].write(str(row.get("Employee", "Unassigned")))
                rcols[7].write(str(row.get("Week", 1)))
                rcols[8].write(f"{int(row['Completion %'])}%")
                rcols[9].write(row["Status"])

                edit_key = f"edit_{idx}"

                # ADMIN-ONLY: Edit, Hide, Delete actions
                if st.session_state.role == "Admin":
                    if rcols[10].button("✏️", key=f"btn_edit_{idx}", help="Edit task"):
                        st.session_state[edit_key] = not st.session_state.get(edit_key, False)

                    # HIDE / UNHIDE
                    hide_label = "👁️" if is_hidden else "🙈"
                    hide_help  = "Unhide task" if is_hidden else "Hide task"
                    if rcols[11].button(hide_label, key=f"btn_hide_{idx}", help=hide_help):
                        df.at[idx, "Hidden"] = not is_hidden
                        save_data(df)
                        st.rerun()

                    # DELETE
                    if rcols[12].button("🗑️", key=f"btn_del_{idx}", help="Delete task"):
                        df = df.drop(index=idx).reset_index(drop=True)
                        save_data(df)
                        st.rerun()

                # INLINE EDITOR — Admin only
                if st.session_state.role == "Admin" and st.session_state.get(edit_key, False):
                    with st.container():
                        st.markdown(f"##### ✏️ Editing Context Component: *{row['Task Name']}*")
                        
                        ec1, ec2, ec3 = st.columns(3)
                        
                        # Project Override
                        safe_proj_idx = projects.index(row["Project"]) if row["Project"] in projects else 0
                        e_proj_sel = ec1.selectbox("Shift Project Category", projects, index=safe_proj_idx, key=f"e_proj_sel_{idx}")
                        e_proj_over = ec1.text_input("OR Rename/Establish New Project", key=f"e_proj_over_{idx}", help="Will override dropdown above")
                        new_proj = e_proj_over.strip() if str(e_proj_over).strip() != "" else e_proj_sel
                        
                        # Topic Override
                        safe_topic_idx = topics.index(row["Topic"]) if row["Topic"] in topics else 0
                        e_topic_sel = ec2.selectbox("Shift Topic Category", topics, index=safe_topic_idx, key=f"e_topic_sel_{idx}")
                        e_topic_over = ec2.text_input("OR Rename/Establish New Topic", key=f"e_topic_over_{idx}", help="Will override dropdown above")
                        new_topic = e_topic_over.strip() if str(e_topic_over).strip() != "" else e_topic_sel
                            
                        # Base Specs
                        new_name   = ec3.text_area("Task Specification", value=row["Task Name"], key=f"e_name_{idx}", height=120)
                        safe_emp_idx = employees.index(row.get("Employee", "Unassigned")) if row.get("Employee", "Unassigned") in employees else 0
                        new_emp = ec3.selectbox("Employee Assigned", employees, index=safe_emp_idx, key=f"e_emp_{idx}")

                        rc1, rc2, rc3, rc4, rc5 = st.columns(5)
                        new_start  = rc1.date_input("Start Configuration", value=pd.to_datetime(row["Start Date"]), key=f"e_start_{idx}")
                        new_end    = rc2.date_input("Close Configuration", value=pd.to_datetime(row["End Date"]), key=f"e_end_{idx}")
                        temp_week_df = df.copy()
                        temp_week_df.at[idx, "Project"] = new_proj
                        temp_week_df.at[idx, "Start Date"] = new_start.strftime("%Y-%m-%d")
                        auto_week = calculate_project_week(new_proj, new_start, temp_week_df)
                        if st.session_state.role == "Admin":
                            new_week = rc3.number_input("Project Week", min_value=1, max_value=2000, value=int(auto_week), step=1, key=f"e_week_{idx}")
                        else:
                            new_week = rc3.number_input("Project Week", min_value=1, max_value=2000, value=int(auto_week), disabled=True, key=f"e_week_{idx}")
                        current_comp = int(row["Completion %"])
                        if st.session_state.role == "Admin":
                            new_comp = rc4.number_input("Done %", min_value=0, max_value=100, value=current_comp, step=1, key=f"e_comp_{idx}")
                        else:
                            # Employee restricted to upward progression only
                            new_comp = rc4.number_input("+% Done", min_value=current_comp, max_value=100, value=current_comp, step=1, key=f"e_comp_{idx}")
                        safe_status_idx = STATUS_OPTIONS.index(row["Status"]) if row["Status"] in STATUS_OPTIONS else 0
                        new_status = rc5.selectbox("Operational Status", STATUS_OPTIONS, index=safe_status_idx, key=f"e_status_{idx}")
                        
                        st.divider()
                        milestone_text = str(row.get("Milestone_Text", ""))
                        milestone_role = str(row.get("Milestone_Role", "None"))
                        milestone_author = str(row.get("Milestone_Author_Name", ""))
                        
                        is_locked_for_emp = (st.session_state.role == "Employee" and milestone_role == "Admin")
                        
                        mc1, mc2 = st.columns([2, 1])
                        if is_locked_for_emp:
                            mc1.info(f"🔒 **Strategy Locked by Administrator** (Last modified by {milestone_author})")
                            new_milestone = mc1.text_input("Milestone Strategy", milestone_text, disabled=True, key=f"e_mile_{idx}")
                        else:
                            if milestone_role != "None":
                                mc1.caption(f"Last modified by **{milestone_author}** ({milestone_role})")
                            new_milestone = mc1.text_input("Milestone Strategy", milestone_text, key=f"e_mile_{idx}")
                        
                        e_files = mc2.file_uploader(f"Add Supplemental Milestone Documents", accept_multiple_files=True, key=f"efiles_{idx}")
                        
                        save_col, cancel_col, _ = st.columns([1, 1, 4])
                        if save_col.button("💾 Apply Overrides", key=f"save_{idx}"):
                            if new_end < new_start:
                                st.error("End Date cannot be earlier than Start Date.")
                            else:
                                df.at[idx, "Project"]      = new_proj
                                df.at[idx, "Topic"]        = new_topic
                                df.at[idx, "Task Name"]    = new_name
                                df.at[idx, "Employee"]     = new_emp
                                df.at[idx, "Start Date"]   = new_start.strftime("%Y-%m-%d")
                                df.at[idx, "End Date"]     = new_end.strftime("%Y-%m-%d")
                                df.at[idx, "Week"]         = new_week
                                df.at[idx, "Completion %"] = new_comp
                                df.at[idx, "Status"]       = new_status
                                
                                if not is_locked_for_emp:
                                    df.at[idx, "Milestone_Text"] = new_milestone
                                    if new_milestone != milestone_text or milestone_role == "None":
                                        df.at[idx, "Milestone_Author_Name"] = st.session_state.auth_name
                                        df.at[idx, "Milestone_Role"] = st.session_state.role
                                        
                                if e_files:
                                    m_dir = f"pmo_storage/{new_proj}/{new_topic}/{new_name}_Milestone"
                                    os.makedirs(m_dir, exist_ok=True)
                                    drive_failures = []
                                    for f in e_files:
                                        _, drive_error = save_uploaded_file(
                                            f,
                                            os.path.join(m_dir, f.name),
                                            [new_proj, new_topic, f"{new_name}_Milestone"]
                                        )
                                        if drive_error:
                                            drive_failures.append(f"{f.name}: {drive_error}")
                                    if drive_failures:
                                        st.warning("Supplemental files saved locally, but some Drive syncs failed: " + " | ".join(drive_failures[:3]))
                                            
                                save_data(df)
                                remember_week_context(new_proj, new_week)
                                
                                st.session_state[edit_key] = False
                                st.success("Network state formally updated!")
                                st.rerun()
                        if cancel_col.button("✖ Abort", key=f"cancel_{idx}"):
                            st.session_state[edit_key] = False
                            st.rerun()
                else:
                    milestone_text = str(row.get("Milestone_Text", "")).strip()
                    if milestone_text:
                        st.markdown(f"**Milestone Strategy:** {format_single_line_text(milestone_text)}")
                        st.markdown("---")

                if st.session_state.get(task_name_toggle_key, False):
                    st.markdown("**Task Name:**")
                    st.markdown(format_bullet_markdown(row["Task Name"]))
                    st.markdown("---")

            # Summary stats
            st.divider()
            total = len(display_df)
            done  = len(display_df[display_df["Status"] == "Completed"])
            delayed = len(display_df[display_df["Status"] == "Delayed"])
            avg_pct = display_df["Completion %"].mean() if total else 0
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("Tasks (in Filter Context)", total)
            s2.metric("Achieved", done)
            s3.metric("Blocked/Delayed", delayed)
            s4.metric("Avg Completion Velocity", f"{avg_pct:.1f}%")

# ─────────────────────────────────────────────
# PLANNED MILESTONES PAGE
# ─────────────────────────────────────────────
elif page == "Planned Milestones":
    st.header("Planned Milestones")
    st.markdown("Structure and coordinate multi-phase milestones, specific tasks, and error tracking.")

    pm_data = load_planned_milestones()

    show_gantt = st.toggle("📊 View Milestone Gantt Chart", value=False)
    if show_gantt:
        gantt_data = []
        for mil_id, mil_info in pm_data.items():
            start = mil_info.get("from_date", "")
            end = mil_info.get("to_date", "")
            if not start or not end: continue
            
            project = mil_info.get("project_context", "Unassigned")
            completed = "Completed" if mil_info.get("completed", False) else "Planned"
            
            gantt_data.append(dict(
                Task_Name=f"🎯 {mil_id}",
                Start=start,
                Finish=end,
                Project=project,
                Status=completed,
                Type="Milestone"
            ))
            
            for t_id, t_info in mil_info.get("tasks", {}).items():
                t_start = t_info.get("from_date", start)
                t_end = t_info.get("to_date", end)
                gantt_data.append(dict(
                    Task_Name=f"└ {t_id} ({mil_id})",
                    Start=t_start,
                    Finish=t_end,
                    Project=project,
                    Status=completed,
                    Type="Task"
                ))
        
        if gantt_data:
            df_gantt = pd.DataFrame(gantt_data)
            df_gantt['Start'] = pd.to_datetime(df_gantt['Start'], errors='coerce')
            df_gantt['Finish'] = pd.to_datetime(df_gantt['Finish'], errors='coerce')
            df_gantt = df_gantt.dropna(subset=['Start', 'Finish'])
            
            if not df_gantt.empty:
                # Add tiny offset if start == finish so bar is visible
                mask_same_date = df_gantt['Start'] == df_gantt['Finish']
                df_gantt.loc[mask_same_date, 'Finish'] += pd.Timedelta(days=1)
                
                df_gantt = df_gantt.sort_values(by=['Project', 'Start'])
                
                fig = px.timeline(
                    df_gantt, 
                    x_start="Start", 
                    x_end="Finish", 
                    y="Task_Name", 
                    color="Status",
                    color_discrete_map={"Completed": "#2ecc71", "Planned": "#3498db"},
                    hover_name="Type",
                    text="Type",
                    title="Planned Milestones & Tasks Timeline"
                )
                fig.update_yaxes(autorange="reversed")
                fig.update_layout(height=max(400, len(df_gantt) * 40))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No valid dates found to generate Gantt chart.")
        else:
            st.info("No milestone data available to generate Gantt chart.")
            
        st.divider()

    # Create root level Add Milestone expander
    with st.expander("➕ Add New Milestone", expanded=False):
        m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns([3, 1, 1, 2, 2])
        new_m_desc = m_col1.text_area("Milestone Description", key="nm_desc")
        new_m_name = m_col2.text_input("Milestone Name", key="nm_name")
        new_m_time = m_col3.number_input("Time Needed (Hours)", min_value=0.0, step=0.5, value=0.0, key="nm_time")
        new_m_project = m_col5.selectbox("Project Context", projects, key="nm_project")
        
        dates_col1, dates_col2 = m_col4.columns(2)
        new_m_start = dates_col1.date_input("From Date", key="nm_start")
        new_m_end = dates_col2.date_input("To Date", key="nm_end")
        nm_proj_topics = get_project_topics(new_m_project, df) if str(new_m_project).strip() != "" else []
        if nm_proj_topics:
            st.markdown("**Topic Progress Increases** *(set how much each topic improves when this milestone is completed)*")
            nm_topic_cols = st.columns(min(len(nm_proj_topics), 4))
            nm_topic_increases = {}
            for ti, tname in enumerate(nm_proj_topics):
                col = nm_topic_cols[ti % 4]
                nm_topic_increases[tname] = col.number_input(
                    f"+% {tname}", min_value=0.0, max_value=100.0, step=1.0, value=0.0,
                    key=f"nm_tinc_{tname}"
                )
        else:
            nm_topic_increases = {}
            st.info("Select a project above to set per-topic progress increases.")

        if st.button("💾 Save Milestone"):
            milestone_name = str(new_m_name).strip()
            if milestone_name == "":
                st.error("Please provide a milestone name.")
            elif milestone_name in pm_data:
                st.error("A milestone with this name already exists. Please choose a different name.")
            elif str(new_m_desc).strip() == "":
                st.error("Please provide a description.")
            elif new_m_end < new_m_start:
                st.error("To Date cannot be earlier than From Date.")
            else:
                pm_data[milestone_name] = {
                    "description": new_m_desc,
                    "time_needed": float(new_m_time),
                    "from_date": new_m_start.strftime("%Y-%m-%d"),
                    "to_date": new_m_end.strftime("%Y-%m-%d"),
                    "project_context": new_m_project,
                    "progress_increase": {k: float(v) for k, v in nm_topic_increases.items()},
                    "completed": False,
                    "tasks": {}
                }
                save_planned_milestones(pm_data)
                
                for k in ["nm_desc", "nm_name", "nm_time", "nm_start", "nm_end", "nm_project"]:
                    if k in st.session_state:
                        del st.session_state[k]

                st.success(f"Added {milestone_name}")
                st.rerun()

    st.divider()
    
    # Display Milestones
    active_milestones = [
        (mil_id, mil_info)
        for mil_id, mil_info in list(pm_data.items())
        if not bool(mil_info.get("completed", False))
    ]
    if not active_milestones:
        st.info("No active planned milestones. Completed milestones now appear on the Tasks & Milestones page.")

    for mil_id, mil_info in active_milestones:
        with st.container():
            header_col1, header_col2, header_col3, header_col4 = st.columns([5, 1.2, 0.9, 0.9])
            header_col1.markdown(f"### 🎯 {mil_id}")
            milestone_done = header_col2.checkbox("Completed", value=False, key=f"done_{mil_id}")
            if milestone_done:
                mil_info["completed"] = True
                save_planned_milestones(pm_data)
                st.rerun()
            
            # --- Edit Mode for Milestone ---
            edit_m_key = f"edit_m_{mil_id}"
            if edit_m_key not in st.session_state:
                st.session_state[edit_m_key] = False

            if st.session_state[edit_m_key]:
                em_c1, em_c2, em_c3, em_c4, em_c5 = st.columns([3, 1, 1, 2, 2])
                e_m_desc = em_c1.text_area("Description", value=mil_info["description"], key=f"emd_{mil_id}")
                e_m_name = em_c2.text_input("Name", value=mil_id, key=f"emn_{mil_id}")
                e_m_time = em_c3.number_input("Time Needed (Hours)", min_value=0.0, step=0.5, value=float(mil_info.get("time_needed", 0)), key=f"emt_{mil_id}")
                current_m_project = str(mil_info.get("project_context", "")).strip()
                project_context_options = projects.copy()
                if current_m_project and current_m_project not in project_context_options:
                    project_context_options.append(current_m_project)
                project_context_index = project_context_options.index(current_m_project) if current_m_project in project_context_options else 0
                e_m_project = em_c5.selectbox("Project Context", project_context_options, index=project_context_index, key=f"emp_{mil_id}")
                
                ed_c1, ed_c2 = em_c4.columns(2)
                cur_start = pd.to_datetime(mil_info.get("from_date", datetime.now().date())).date()
                cur_end = pd.to_datetime(mil_info.get("to_date", datetime.now().date())).date()
                e_m_start = ed_c1.date_input("From Date", value=cur_start, key=f"ems_{mil_id}")
                e_m_end = ed_c2.date_input("To Date", value=cur_end, key=f"eme_{mil_id}")
                em_proj_topics = get_project_topics(e_m_project, df) if str(e_m_project).strip() != "" else []
                existing_increases = get_milestone_topic_increases(mil_info)
                if em_proj_topics:
                    st.markdown("**Topic Progress Increases**")
                    em_topic_cols = st.columns(min(len(em_proj_topics), 4))
                    e_topic_increases = {}
                    for ti, tname in enumerate(em_proj_topics):
                        col = em_topic_cols[ti % 4]
                        e_topic_increases[tname] = col.number_input(
                            f"+% {tname}", min_value=0.0, max_value=100.0, step=1.0,
                            value=float(existing_increases.get(tname, 0.0)),
                            key=f"em_tinc_{mil_id}_{tname}"
                        )
                else:
                    e_topic_increases = existing_increases
                
                scol1, scol2 = st.columns([1, 10])
                if scol1.button("💾", key=f"sm_{mil_id}"):
                    new_milestone_name = str(e_m_name).strip()
                    if new_milestone_name == "":
                        st.error("Please provide a milestone name.")
                    elif new_milestone_name != mil_id and new_milestone_name in pm_data:
                        st.error("A milestone with this name already exists. Please choose a different name.")
                    elif e_m_end < e_m_start:
                        st.error("To Date cannot be earlier than From Date.")
                    else:
                        if new_milestone_name != mil_id:
                            pm_data[new_milestone_name] = pm_data.pop(mil_id)
                            mil_info = pm_data[new_milestone_name]
                            st.session_state[f"edit_m_{new_milestone_name}"] = st.session_state.pop(edit_m_key, False)
                            edit_m_key = f"edit_m_{new_milestone_name}"
                            mil_id = new_milestone_name
                        mil_info["description"] = str(e_m_desc)
                        mil_info["time_needed"] = float(e_m_time)
                        mil_info["from_date"] = e_m_start.strftime("%Y-%m-%d")
                        mil_info["to_date"] = e_m_end.strftime("%Y-%m-%d")
                        mil_info["project_context"] = e_m_project
                        mil_info["progress_increase"] = {k: float(v) for k, v in e_topic_increases.items()}
                        if "topic" in mil_info:
                            del mil_info["topic"]
                        save_planned_milestones(pm_data)
                        st.session_state[edit_m_key] = False
                        st.rerun()
                if scol2.button("❌", key=f"cm_{mil_id}"):
                    st.session_state[edit_m_key] = False
                    st.rerun()
            else:
                md_c1, md_c2, md_c3, md_c4 = st.columns([4.4, 2.2, 0.8, 0.8])
                md_c1.markdown("**Description:**")
                md_c1.markdown(format_bullet_markdown(mil_info.get("description", "")))
                topic_inc_display = get_milestone_topic_increases(mil_info)
                inc_lines = "  \n".join([f"**+{v:.0f}%** {t}" for t, v in topic_inc_display.items() if v > 0]) or "No increase set"
                md_c2.markdown(
                    f"**Project:** {mil_info.get('project_context', 'Not linked')}  \n"
                    f"**Time:** {mil_info.get('time_needed', 0)} Hrs  \n"
                    f"**Dates:** {mil_info.get('from_date', '')} to {mil_info.get('to_date', '')}  \n"
                    f"**Topic Increases:**  \n{inc_lines}"
                )
                if md_c3.button("✏️ Edit", key=f"bem_{mil_id}"):
                    st.session_state[edit_m_key] = True
                    st.rerun()
                if md_c4.button("🗑️ Del", key=f"delm_{mil_id}"):
                    del pm_data[mil_id]
                    save_planned_milestones(pm_data)
                    st.rerun()

            # Tasks Display
            m_tasks = mil_info.get("tasks", {})
            with st.expander(f"#### Tasks {mil_id}", expanded=False):
                if not m_tasks:
                    st.info("No tasks added yet for this milestone.")
                for t_id, t_info in list(m_tasks.items()):
                    with st.container():
                        row_col1, row_col2, row_col3 = st.columns([4.2, 1.5, 1.8])
                        row_col1.markdown(f"**{t_id}**")
                        row_col1.markdown(format_bullet_markdown(t_info.get("description", "")))
                        row_col2.markdown(f"⏱️ {t_info.get('time_needed', 0)} Hrs")
                        row_col2.markdown(f"📅 {t_info.get('from_date', '')} - {t_info.get('to_date','')}")
                        completed_key = f"task_done_{mil_id}_{t_id}"
                        current_completed = bool(t_info.get("completed", False))
                        task_completed = row_col3.checkbox("Completed", value=current_completed, key=completed_key)
                        if task_completed != current_completed:
                            t_info["completed"] = task_completed
                            save_planned_milestones(pm_data)
                            st.rerun()
                        action_col1, action_col2 = row_col3.columns(2)
                        
                        edit_t_key = f"edit_t_{mil_id}_{t_id}"
                        if edit_t_key not in st.session_state:
                             st.session_state[edit_t_key] = False
                             
                        if action_col1.button("✏️", key=f"bet_{mil_id}_{t_id}"):
                             st.session_state[edit_t_key] = not st.session_state[edit_t_key]
                             
                        if action_col2.button("🗑️", key=f"delt_{mil_id}_{t_id}"):
                             del m_tasks[t_id]
                             save_planned_milestones(pm_data)
                             st.rerun()
                             
                        if st.session_state.get(edit_t_key, False):
                            et_c1, et_c2, et_c3, et_c4 = st.columns([3, 1, 1, 2])
                            et_desc = et_c1.text_area("Task Description", value=t_info["description"], key=f"etd_{mil_id}_{t_id}")
                            et_name = et_c2.text_input("Name", value=t_id, key=f"etn_{mil_id}_{t_id}")
                            et_time = et_c3.number_input("Time Needed", min_value=0.0, step=0.5, value=float(t_info.get("time_needed", 0)), key=f"ett_{mil_id}_{t_id}")
                            
                            cur_m_start = pd.to_datetime(mil_info.get("from_date", datetime.now().date())).date()
                            cur_m_end = pd.to_datetime(mil_info.get("to_date", datetime.now().date())).date()
                            try:
                                # Safely handle existing task dates if they are unexpectedly outside bounds
                                t_start_val = pd.to_datetime(t_info.get("from_date", cur_m_start)).date()
                                t_end_val = pd.to_datetime(t_info.get("to_date", cur_m_end)).date()
                                t_start_val = max(cur_m_start, min(cur_m_end, t_start_val))
                                t_end_val = max(cur_m_start, min(cur_m_end, t_end_val))
                            except:
                                t_start_val = cur_m_start
                                t_end_val = cur_m_end

                            etd_c1, etd_c2 = et_c4.columns(2)
                            et_s = etd_c1.date_input("From Date", value=t_start_val, min_value=cur_m_start, max_value=cur_m_end, key=f"ets_{mil_id}_{t_id}")
                            et_e = etd_c2.date_input("To Date", value=t_end_val, min_value=cur_m_start, max_value=cur_m_end, key=f"ete_{mil_id}_{t_id}")
                            
                            esc1, esc2, _ = st.columns([1, 1, 10])
                            if esc1.button("💾", key=f"est_{mil_id}_{t_id}"):
                                new_task_name = str(et_name).strip()
                                if new_task_name == "":
                                    st.error("Please provide a task name.")
                                elif new_task_name != t_id and new_task_name in m_tasks:
                                    st.error("A task with this name already exists in this milestone.")
                                else:
                                    if new_task_name != t_id:
                                        m_tasks[new_task_name] = m_tasks.pop(t_id)
                                        t_info = m_tasks[new_task_name]
                                        st.session_state[f"edit_t_{mil_id}_{new_task_name}"] = st.session_state.pop(edit_t_key, False)
                                        edit_t_key = f"edit_t_{mil_id}_{new_task_name}"
                                        t_id = new_task_name
                                    t_info["description"] = str(et_desc)
                                    t_info["time_needed"] = float(et_time)
                                    t_info["from_date"] = et_s.strftime("%Y-%m-%d")
                                    t_info["to_date"] = et_e.strftime("%Y-%m-%d")
                                    save_planned_milestones(pm_data)
                                    st.session_state[edit_t_key] = False
                                    st.rerun()
                            if esc2.button("❌", key=f"cst_{mil_id}_{t_id}"):
                                st.session_state[edit_t_key] = False
                                st.rerun()

                        impact_project = str(mil_info.get("project_context", t_info.get("project", ""))).strip()
                        current_topic = str(t_info.get("topic", "")).strip()
                        topic_label = current_topic if current_topic else "Not linked"
                        project_label = impact_project if impact_project else "Not linked"
                        row_col1.caption(
                            f"Task Link: {project_label} / {topic_label}"
                        )
                        
                        # Display errors inside task
                        if "errors" in t_info and len(t_info["errors"]) > 0:
                            for idx, err in enumerate(t_info["errors"]):
                                err_cols = st.columns([0.5, 8, 1])
                                timing_warn = " ⚠️ *(Timing May Vary)*" if err.get("timing_varied", False) else ""
                                err_cols[1].caption(f"⚠️ **Error/New Task {idx+1}:** {err['description']} *(+ {err['hours_spent']} hrs)*{timing_warn}")
                                solution_text = str(err.get("solution", "")).strip()
                                if solution_text:
                                    err_cols[1].markdown("**Solution / Fix:**")
                                    err_cols[1].markdown(format_bullet_markdown(solution_text))
                                if err_cols[2].button("🗑️", key=f"delerr_{mil_id}_{t_id}_{idx}"):
                                    t_info["time_needed"] = float(t_info.get("time_needed", 0)) - err["hours_spent"]
                                    mil_info["time_needed"] = float(mil_info.get("time_needed", 0)) - err["hours_spent"]
                                    t_info["errors"].pop(idx)
                                    save_planned_milestones(pm_data)
                                    st.rerun()
            
            # --- Add Task inside Milestone ---
            with st.expander(f"➕ Add Task to {mil_id}", expanded=False):
                milestone_project = str(mil_info.get("project_context", "")).strip()
                t_c1, t_c2, t_c3, t_c5 = st.columns([2, 3, 1, 3])
                new_t_name = t_c1.text_input("Task Name", key=f"ntname_{mil_id}")
                new_t_desc = t_c2.text_area("Task Description", key=f"ntd_{mil_id}")
                new_t_time = t_c3.number_input("Time Needed (Total)", min_value=0.0, step=0.5, value=0.0, key=f"ntt_{mil_id}")
                
                t_d1, t_d2 = t_c5.columns(2)
                cur_m_start = pd.to_datetime(mil_info.get("from_date", datetime.now().date())).date()
                cur_m_end = pd.to_datetime(mil_info.get("to_date", datetime.now().date())).date()
                
                new_t_start = t_d1.date_input("From Date", value=cur_m_start, min_value=cur_m_start, max_value=cur_m_end, key=f"nts_{mil_id}")
                new_t_end = t_d2.date_input("To Date", value=cur_m_end, min_value=cur_m_start, max_value=cur_m_end, key=f"nte_{mil_id}")

                if st.button("💾 Save Task(s)", key=f"st_{mil_id}"):
                    task_name = str(new_t_name).strip()
                    if task_name == "":
                        st.error("Please provide a task name.")
                    elif task_name in m_tasks:
                        st.error("A task with this name already exists in this milestone.")
                    elif str(new_t_desc).strip() == "":
                        st.error("Please provide a task description.")
                    else:
                        m_tasks[task_name] = {
                            "description": str(new_t_desc),
                            "time_needed": float(new_t_time),
                            "from_date": new_t_start.strftime("%Y-%m-%d"),
                            "to_date": new_t_end.strftime("%Y-%m-%d"),
                            "project": milestone_project,
                            "topic": "",
                            "completed": False,
                            "errors": []
                        }
                        mil_info["tasks"] = m_tasks
                        save_planned_milestones(pm_data)
                        
                        for k in [f"ntname_{mil_id}", f"ntd_{mil_id}", f"ntt_{mil_id}", f"nts_{mil_id}", f"nte_{mil_id}"]:
                            if k in st.session_state:
                                del st.session_state[k]

                        st.success(f"Added {task_name} to {mil_id}")
                        st.rerun()

            # --- Errors / New Tasks ---
            with st.expander(f"⚠️ Report Errors / Additional Tasks for {mil_id}", expanded=False):
                if not m_tasks:
                    st.info("No tasks exist yet in this milestone to assign errors to.")
                else:
                    err_c1, err_c2, err_c3, err_c4 = st.columns([2, 4, 1.5, 1.5])
                    err_task = err_c1.selectbox("Target Task", options=list(m_tasks.keys()), key=f"errt_{mil_id}")
                    err_desc = err_c2.text_area("Error / New Task ", key=f"errd_{mil_id}", height=90)
                    err_hrs = err_c3.number_input("Hours Spent", min_value=0.0, step=0.5, value=1.0, key=f"errh_{mil_id}")
                    err_timing = err_c4.checkbox("Timing May Vary", value=False, key=f"errtm_{mil_id}", help="Check if milestone timing may vary because of this error/task")
                    err_solution = st.text_area("Description", key=f"errs_{mil_id}", height=90)
                    
                    if st.button("💾 Log Issue & Update Times", key=f"b_err_{mil_id}"):
                        if str(err_desc).strip() != "":
                            if "errors" not in m_tasks[err_task]:
                                m_tasks[err_task]["errors"] = []
                            m_tasks[err_task]["errors"].append({
                                "description": err_desc,
                                "solution": err_solution,
                                "hours_spent": float(err_hrs),
                                "timing_varied": err_timing
                            })
                            # Auto-update times
                            m_tasks[err_task]["time_needed"] = float(m_tasks[err_task].get("time_needed", 0)) + float(err_hrs)
                            mil_info["time_needed"] = float(mil_info.get("time_needed", 0)) + float(err_hrs)
                            
                            save_planned_milestones(pm_data)
                            
                            for k in [f"errt_{mil_id}", f"errd_{mil_id}", f"errs_{mil_id}", f"errh_{mil_id}", f"errtm_{mil_id}"]:
                                if k in st.session_state:
                                    del st.session_state[k]

                            st.success(f"Logged issue to {err_task}. {err_hrs} hrs added to task and milestone totals.")
                            st.rerun()
                        else:
                            st.error("Please provide an issue description.")
            st.divider()


# ─────────────────────────────────────────────
# IMAGE GALLERY PAGE
# ─────────────────────────────────────────────
elif page == "Image Gallery":
    st.header("🖼️ Dashboard Image Gallery")
    st.markdown("Manage the subsystem images that automatically display below dashboard gauges.")
    
    os.makedirs("data/topic_images", exist_ok=True)
    
    # Upload Form
    with st.expander("➕ Link New Image to Dashboard Target", expanded=True):
        fcol1, fcol2 = st.columns(2)
        up_proj = fcol1.selectbox("Target Project", projects)
        up_topic = fcol2.selectbox("Target Subsystem/Topic", topics)
        uploaded = st.file_uploader("Upload Image Payload", type=["png", "jpg", "jpeg", "webp"])
        
        if st.button("💾 Save to Graphic Engine Directory"):
            if uploaded:
                img_path = f"data/topic_images/{up_proj}_{up_topic}.png"
                _, drive_error = save_uploaded_file(
                    uploaded,
                    img_path,
                    [up_proj, up_topic, "_topic_images"]
                )
                st.session_state.last_drive_status = {
                    "project": up_proj,
                    "topic": up_topic,
                    "drive_error": drive_error,
                    "saved_to_drive": drive_error is None,
                }
                st.success(f"Success! Image locked to '{up_proj} : {up_topic}'.")
                if drive_error:
                    st.warning(f"Image saved locally, but Google Drive sync failed: {drive_error}")
                st.rerun()
            else:
                st.error("Please load a file first.")

    if st.session_state.last_drive_status:
        last_status = st.session_state.last_drive_status
        if last_status["saved_to_drive"]:
            st.info(
                f"Last image upload for {last_status['project']} : {last_status['topic']} synced to Google Drive."
            )
        else:
            st.warning(
                f"Last image upload for {last_status['project']} : {last_status['topic']} was only saved locally. "
                f"Drive error: {last_status['drive_error']}"
            )

    st.divider()
    st.subheader("Currently Active Directory Images")
    
    existing_images = sorted([f for f in os.listdir("data/topic_images") if f.endswith(".png")])
    if existing_images:
        gcols = st.columns(3)
        for i, img_file in enumerate(existing_images):
            # Resolve original config from filename if possible
            caption_str = img_file.replace(".png", "").replace("_", " : ")
            with gcols[i % 3]:
                st.markdown(f"**{caption_str}**")
                st.image(f"data/topic_images/{img_file}", use_container_width=True)
                if st.button(f"🗑️ Erase Link", key=f"del_{img_file}"):
                    os.remove(f"data/topic_images/{img_file}")
                    st.rerun()
                st.markdown("---")
    else:
        st.info("The graphics directory is currently empty. Upload images above to populate gauges.")

# ─────────────────────────────────────────────
# COMPETITORS PAGE
# ─────────────────────────────────────────────
elif page == "Competitors & Research":
    st.header("Competitor Benchmarking")
    comp_data = load_competitor_data()

    if st.session_state.role == "Admin":
        with st.expander("➕ Add New Benchmark Topic", expanded=False):
            new_topic = st.text_input("New Topic Name (e.g. SLAM Accuracy)")
            new_cols = st.text_input("Column Names (comma separated)", value="Competitor, Value, Notes", help="Define columns for your new table")
            if st.button("Create Topic"):
                if new_topic and new_topic not in comp_data:
                    c_list = [c.strip() for c in new_cols.split(",") if c.strip()]
                    if not c_list:
                        c_list = ["Competitor", "Value"]
                    empty_row = {c: "" for c in c_list}
                    comp_data[new_topic] = [empty_row]
                    save_competitor_data(comp_data)
                    st.success(f"Topic '{new_topic}' created.")
                    st.rerun()
                elif new_topic in comp_data:
                    st.error("Topic already exists.")

    for topic, rows in list(comp_data.items()):
        if rows:
            columns = list(rows[0].keys())
        else:
            columns = ["Competitor", "Value"]

        with st.expander(f"📊 {topic}", expanded=False):
            if st.session_state.role == "Admin":
                st.markdown("Use the table below to **Edit Data**, **Add Rows**, or **Delete Rows** directly. Click '💾 Save Table Edits' to apply.")
                
                has_substance = bool(rows and any(any(str(v).strip() for v in r.values()) for r in rows))
                if has_substance:
                    df_topic = pd.DataFrame(rows)
                else:
                    df_topic = pd.DataFrame(columns=columns)
                
                edited_df = st.data_editor(df_topic, num_rows="dynamic", key=f"editor_{topic}", use_container_width=True)
                
                if st.button("💾 Save Table Edits", key=f"save_edits_{topic}"):
                    comp_data[topic] = edited_df.fillna("").to_dict(orient="records")
                    if not comp_data[topic]:
                         comp_data[topic] = [{c: "" for c in columns}]
                    save_competitor_data(comp_data)
                    st.success("Table edits saved!")
                    st.rerun()

                st.markdown("---")
                st.markdown("**Column Management & Admin Actions**")

                col_add, col_del = st.columns(2)
                with col_add:
                    st.markdown("##### ➕ Add Column")
                    new_col_name = st.text_input("New Column Name", key=f"new_col_name_{topic}")
                    if st.button("Add Column", key=f"add_col_btn_{topic}"):
                        if new_col_name and new_col_name not in columns:
                            # Use current state from editor instead of stale comp_data
                            current_rows = edited_df.fillna("").to_dict(orient="records")
                            if not current_rows:
                                current_rows = [{c: "" for c in columns}]
                            
                            for r in current_rows:
                                r[new_col_name] = ""
                            
                            comp_data[topic] = current_rows
                            save_competitor_data(comp_data)
                            st.rerun()

                with col_del:
                    st.markdown("##### ❌ Delete Column")
                    col_to_delete = st.selectbox("Select Column to Delete", columns, key=f"del_col_sel_{topic}")
                    if st.button("Delete Column", key=f"del_col_btn_{topic}"):
                        if len(columns) <= 1:
                            st.error("Cannot delete the last column.")
                        else:
                            # Use current state from editor
                            current_rows = edited_df.fillna("").to_dict(orient="records")
                            for r in current_rows:
                                if col_to_delete in r:
                                    del r[col_to_delete]
                            
                            comp_data[topic] = current_rows
                            save_competitor_data(comp_data)
                            st.rerun()

                st.divider()
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button(f"🗑️ Delete Topic '{topic}'", key=f"del_topic_{topic}", type="primary"):
                    del comp_data[topic]
                    save_competitor_data(comp_data)
                    st.rerun()
            else:
                if rows:
                    display_rows = [r for r in rows if any(str(v).strip() for v in r.values())]
                    if display_rows:
                        df_topic = pd.DataFrame(display_rows)
                        st.table(df_topic)
                    else:
                        st.info("No data added for this topic yet (Template ready).")
                else:
                    st.info("No data added for this topic yet.")



# ─────────────────────────────────────────────
# DOCUMENT DRIVE PAGE
# ─────────────────────────────────────────────
elif page == "Document Drive":
    st.header("Secure Storage & Topic Targeting")
    
    if st.session_state.role != "Admin":
        st.error("🔒 Access Denied: Administrator clearance required to access the Document Drive.")
        st.stop()
        
    st.info("Files uploaded here will be securely anchored to specific Projects & Topics!")

    d_col1, d_col2 = st.columns(2)
    t_proj = d_col1.selectbox("Target Project Context", projects)
    t_topic = d_col2.selectbox("Target Subsystem/Topic", topics)

    drive_meta = load_drive_metadata()
    if t_proj not in drive_meta:
        drive_meta[t_proj] = {}
    if t_topic not in drive_meta[t_proj]:
        drive_meta[t_proj][t_topic] = {"file_notes": {}, "urls": []}
    
    topic_meta = drive_meta[t_proj][t_topic]

    st.markdown("### 1. Upload Files")
    uploaded_files = st.file_uploader(f"Upload Data to {t_topic}", accept_multiple_files=True)

    if uploaded_files:
        if st.button("⬆️ Upload Files to Drive", key="doc_drive_upload_btn"):
            target_dir = f"pmo_storage/{t_proj}/{t_topic}"
            os.makedirs(target_dir, exist_ok=True)
            drive_failures = []
            for f in uploaded_files:
                _, drive_error = save_uploaded_file(
                    f,
                    os.path.join(target_dir, f.name),
                    [t_proj, t_topic]
                )
                if drive_error:
                    drive_failures.append(f"{f.name}: {drive_error}")
            if drive_failures:
                st.warning("Some files were saved locally but not synced to Google Drive: " + " | ".join(drive_failures[:3]))
            else:
                st.success(f"✅ Successfully uploaded {len(uploaded_files)} file(s) to {t_proj} > {t_topic} on Google Drive.")
            st.rerun()

    st.markdown("### 2. Attach URL / Drive Link")
    u_col1, u_col2 = st.columns([2, 3])
    new_url = u_col1.text_input("Reference URL Hook", key="new_url", placeholder="e.g. https://drive.google.com/...")
    new_url_note = u_col2.text_input("Small Note / Description", key="new_url_note", placeholder="What does this link refer to?")
    if st.button("Attach Link"):
        if str(new_url).strip() != "":
            topic_meta["urls"].append({"url": new_url, "note": new_url_note})
            save_drive_metadata(drive_meta)
            st.success("Link attached!")
            st.rerun()
        else:
            st.error("Please enter a URL first.")

    st.divider()

    target_dir = f"pmo_storage/{t_proj}/{t_topic}"
    meta_path = os.path.join(target_dir, ".metadata.json")
    topic_dir_meta = {}
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r") as mf:
                topic_dir_meta = json.load(mf)
        except:
            pass

    st.markdown(f"### 📂 Assets for **{t_proj}** » **{t_topic}**")
    
    with st.expander("🔗 View Uploaded Links", expanded=False):
        has_links = False
        if len(topic_meta["urls"]) > 0:
            has_links = True
            for idx, url_entry in enumerate(topic_meta["urls"]):
                with st.container():
                    c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
                    c1.markdown(f"[{url_entry['url']}]({url_entry['url']})")
                    updated_note = c2.text_input("Note", value=url_entry.get("note", ""), key=f"unote_{idx}", label_visibility="collapsed")
                    if c3.button("💾 Save", key=f"su_{idx}"):
                        topic_meta["urls"][idx]["note"] = updated_note
                        save_drive_metadata(drive_meta)
                        st.success("Saved!")
                    if c4.button("🗑️ Del", key=f"du_{idx}"):
                        topic_meta["urls"].pop(idx)
                        save_drive_metadata(drive_meta)
                        st.rerun()

        dir_links = topic_dir_meta.get("links", {})
        if dir_links:
            has_links = True
            for link_id, link_info in list(dir_links.items()):
                with st.container():
                    c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
                    l_url = link_info.get("url", "")
                    c1.markdown(f"[{l_url}]({l_url})")
                    updated_note = c2.text_input("Note", value=link_info.get("note", ""), key=f"unote_dir_{link_id}", label_visibility="collapsed")
                    if c3.button("💾 Save", key=f"su_dir_{link_id}"):
                        topic_dir_meta["links"][link_id]["note"] = updated_note
                        with open(meta_path, "w") as mf:
                            json.dump(topic_dir_meta, mf, indent=4)
                        st.success("Saved!")
                    if c4.button("🗑️ Del", key=f"du_dir_{link_id}"):
                        del topic_dir_meta["links"][link_id]
                        with open(meta_path, "w") as mf:
                            json.dump(topic_dir_meta, mf, indent=4)
                        st.rerun()

        if not has_links:
            st.info("No links attached yet.")
        
    with st.expander("📄 View Uploaded Files", expanded=False):
        files = []
        if os.path.exists(target_dir):
            files = [f for f in os.listdir(target_dir) if f != ".metadata.json"]
        meta_files = topic_dir_meta.get("files", {})
        for f_name in meta_files.keys():
            if f_name not in files:
                files.append(f_name)
                
        if files:
            for idx, file_item in enumerate(files):
                with st.container():
                    c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
                    file_path = os.path.join(target_dir, file_item)
                    d_url = meta_files.get(file_item, {}).get("drive_url", "")
                    
                    if os.path.exists(file_path):
                        with open(file_path, "rb") as bf:
                            c1.download_button(label=f"⬇️ {file_item}", data=bf, file_name=file_item, mime="application/octet-stream", key=f"dl_{idx}")
                    else:
                        if d_url:
                            c1.markdown(f"☁️ [View file on Drive: {file_item}]({d_url})")
                        else:
                            c1.warning(f"⚠️ {file_item} (Missing)")
                    
                    current_fnote = topic_meta.get("file_notes", {}).get(file_item, "")
                    if not current_fnote and "files" in topic_dir_meta and file_item in topic_dir_meta["files"]:
                        current_fnote = topic_dir_meta["files"][file_item].get("note", "")

                    updated_fnote = c2.text_input("Note", value=current_fnote, key=f"fnote_{idx}", label_visibility="collapsed", placeholder="Add a short note...")
                    if c3.button("💾 Save", key=f"sf_{idx}"):
                        if "file_notes" not in topic_meta:
                            topic_meta["file_notes"] = {}
                        topic_meta["file_notes"][file_item] = updated_fnote
                        
                        if "files" not in topic_dir_meta:
                            topic_dir_meta["files"] = {}
                        if file_item not in topic_dir_meta["files"]:
                            topic_dir_meta["files"][file_item] = {}
                        topic_dir_meta["files"][file_item]["note"] = updated_fnote
                        
                        save_drive_metadata(drive_meta)
                        with open(meta_path, "w") as mf:
                            json.dump(topic_dir_meta, mf, indent=4)
                        st.success("Saved!")
                    if c4.button("🗑️ Del", key=f"df_{idx}"):
                        if os.path.exists(file_path):
                            try:
                                os.remove(file_path)
                            except:
                                pass
                        if file_item in topic_meta.get("file_notes", {}):
                            del topic_meta["file_notes"][file_item]
                            save_drive_metadata(drive_meta)
                        if "files" in topic_dir_meta and file_item in topic_dir_meta["files"]:
                            del topic_dir_meta["files"][file_item]
                            with open(meta_path, "w") as mf:
                                json.dump(topic_dir_meta, mf, indent=4)
                        st.rerun()
        else:
            st.info("No files uploaded for this topic yet.")



