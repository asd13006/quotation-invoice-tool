"""工程項目追蹤 — JSON file store for project lifecycle, status, payments, dashboard stats"""
import json
import os
import uuid
from datetime import datetime, timezone

_PROJECTS_FILE = os.path.join(os.path.dirname(__file__), 'projects.json')

STATUSES = ['draft', 'sent', 'confirmed', 'in_progress', 'completed', 'cancelled']
STATUS_LABELS = {
    'draft': '草稿', 'sent': '已發出', 'confirmed': '已確認',
    'in_progress': '施工中', 'completed': '已完工', 'cancelled': '已取消',
}


def _load():
    if _is_vercel():
        return _load_from_drive()
    if not os.path.exists(_PROJECTS_FILE):
        return {'projects': [], 'counter_quotation': 0, 'counter_invoice': 0}
    with open(_PROJECTS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def _save(db):
    with open(_PROJECTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(db, f, ensure_ascii=False, indent=2)
    if _is_vercel():
        _save_to_drive(db)

def _is_vercel():
    return bool(os.environ.get('VERCEL'))

_DRIVE_DB_NAME = '_project_tracker.json'
_drive_db_id = None

def _load_from_drive():
    """Download projects.json from Google Drive"""
    global _drive_db_id
    try:
        from drive_sync import _get_drive
        drive = _get_drive()
        # Search for existing db file
        results = drive.files().list(
            q=f"name='{_DRIVE_DB_NAME}' and trashed=false",
            pageSize=1, fields='files(id)'
        ).execute()
        files = results.get('files', [])
        if files:
            _drive_db_id = files[0]['id']
            req = drive.files().get_media(fileId=_drive_db_id)
            data = json.loads(req.execute().decode('utf-8'))
            return data
    except Exception:
        pass
    return {'projects': [], 'counter_quotation': 0, 'counter_invoice': 0}

def _save_to_drive(db):
    """Upload projects.json to Google Drive"""
    global _drive_db_id
    try:
        from drive_sync import _get_drive
        from googleapiclient.http import MediaIoBaseUpload
        drive = _get_drive()
        content = json.dumps(db, ensure_ascii=False, indent=2).encode('utf-8')
        media = MediaIoBaseUpload(io.BytesIO(content), mimetype='application/json', resumable=True)
        if _drive_db_id:
            drive.files().update(fileId=_drive_db_id, media_body=media).execute()
        else:
            f = drive.files().create(
                body={'name': _DRIVE_DB_NAME, 'mimeType': 'application/json'},
                media_body=media, fields='id'
            ).execute()
            _drive_db_id = f.get('id')
    except Exception as e:
        print(f'Drive save failed: {e}')



def _now_iso():
    return datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')


def _next_number(db, prefix):
    year = datetime.now().strftime('%Y')
    key = 'counter_' + prefix
    db[key] = db.get(key, 0) + 1
    return prefix.upper() + '-' + year + '-' + str(db[key]).zfill(4)


def register_project(data, doc_type='quotation'):
    db = _load()

    proj_name = (data.get('project_name') or data.get('address') or '未命名').strip()
    if not proj_name:
        proj_name = '未命名工程'

    existing = None
    addr = (data.get('address') or '').strip()
    qno = (data.get('quotation_no') or '').strip()
    for p in db['projects']:
        if addr and p.get('address') == addr:
            existing = p
            break
        if qno and p.get('quotation_no') == qno:
            existing = p
            break

    now = _now_iso()
    items = data.get('items', [])
    subtotal = sum((it.get('quantity', 1) or 1) * (it.get('unit_price', 0) or 0) for it in items)
    deposit = int(data.get('deposit', 0)) or 0

    if existing:
        existing['updated_at'] = now
        existing['status'] = existing.get('status', 'draft')
        if doc_type == 'invoice':
            existing['invoice_no'] = qno
            existing['invoice_date'] = data.get('date', '')
            existing['invoice_total'] = subtotal
            existing['invoice_deposit'] = deposit
            existing['status'] = 'confirmed'
            if 'payments_paid' not in existing:
                existing['payments_paid'] = [False] * len(data.get('payments', []))
        else:
            existing['quotation_no'] = qno
            existing['quotation_date'] = data.get('date', '')
            existing['quotation_total'] = subtotal
        existing['items'] = items
        existing['payments'] = data.get('payments', [])
        existing['terms'] = data.get('terms', [])
    else:
        proj = {
            'id': uuid.uuid4().hex[:12],
            'project_name': proj_name,
            'owner_name': (data.get('owner_name') or '').strip(),
            'address': addr,
            'company_name': (data.get('company_name') or '').strip(),
            'status': 'draft',
            'created_at': now,
            'updated_at': now,
            'quotation_no': qno if doc_type == 'quotation' else '',
            'quotation_date': data.get('date', '') if doc_type == 'quotation' else '',
            'quotation_total': subtotal if doc_type == 'quotation' else 0,
            'invoice_no': qno if doc_type == 'invoice' else '',
            'invoice_date': data.get('date', '') if doc_type == 'invoice' else '',
            'invoice_total': subtotal if doc_type == 'invoice' else 0,
            'invoice_deposit': deposit if doc_type == 'invoice' else 0,
            'total': subtotal,
            'deposit': deposit,
            'items': items,
            'payments': data.get('payments', []),
            'terms': data.get('terms', []),
            'payments_paid': [False] * len(data.get('payments', [])),
            'show_payment': data.get('show_payment', True),
            'show_terms': data.get('show_terms', True),
        }
        db['projects'].append(proj)
        existing = proj

    _save(db)
    return existing


def update_status(project_id, new_status):
    db = _load()
    for p in db['projects']:
        if p['id'] == project_id:
            p['status'] = new_status
            p['updated_at'] = _now_iso()
            _save(db)
            return p
    return None


def toggle_payment(project_id, payment_index):
    db = _load()
    for p in db['projects']:
        if p['id'] == project_id:
            paid = p.get('payments_paid', [])
            if payment_index < len(paid):
                paid[payment_index] = not paid[payment_index]
                p['payments_paid'] = paid
                p['updated_at'] = _now_iso()
                _save(db)
                return p
    return None


def list_projects_local(status_filter=None):
    db = _load()
    projects = db['projects']
    if status_filter:
        projects = [p for p in projects if p.get('status') == status_filter]
    return sorted(projects, key=lambda p: p.get('updated_at', ''), reverse=True)


def get_project(project_id):
    db = _load()
    for p in db['projects']:
        if p['id'] == project_id:
            return p
    return None


def delete_project(project_id):
    db = _load()
    for p in db['projects']:
        if p['id'] == project_id:
            p['status'] = 'cancelled'
            p['updated_at'] = _now_iso()
            _save(db)
            return p
    return None




def sync_from_drive(drive_files):
    """Sync Google Drive file metadata into projects.json.
    Only adds new files; does not overwrite existing projects.
    """
    db = _load()
    existing_ids = {p.get('id', '') for p in db['projects']}
    existing_names = {p.get('project_name', '') for p in db['projects']}
    added = 0
    now = _now_iso()
    for df in drive_files:
        fid = df.get('id', '')
        name = df.get('name', '')
        if fid and fid in existing_ids:
            continue
        if name and name in existing_names:
            continue
        proj = {
            'id': fid,
            'project_name': name,
            'owner_name': '',
            'address': '',
            'company_name': '',
            'status': 'draft',
            'created_at': df.get('date', '') + 'T00:00:00Z' if df.get('date') else now,
            'updated_at': now,
            'quotation_no': '',
            'quotation_date': df.get('date', ''),
            'quotation_total': 0,
            'invoice_no': '',
            'invoice_date': '',
            'invoice_total': 0,
            'invoice_deposit': 0,
            'total': 0,
            'deposit': 0,
            'items': [],
            'payments': [],
            'terms': [],
            'payments_paid': [],
            'show_payment': False,
            'show_terms': False,
            'drive_name': name,
            'drive_size_kb': df.get('size_kb', 0),
            'drive_folder_id': df.get('folder_id', ''),
        }
        db['projects'].append(proj)
        existing_ids.add(fid)
        existing_names.add(name)
        added += 1
    if added > 0:
        _save(db)
    return added



def parse_drive_files(limit=10):
    """Download and parse Drive xlsx files that haven't been parsed yet.
    Returns number of files parsed."""
    db = _load()
    parsed = 0
    for p in db['projects']:
        if parsed >= limit:
            break
        # Skip if already has items (already parsed)
        if p.get('items') and len(p.get('items', [])) > 0:
            continue
        fid = p.get('id', '')
        if not fid or len(fid) < 10:
            continue
        try:
            from drive_sync import get_project_data
            data = get_project_data(fid)
            if data.get('items'):
                p['items'] = data['items']
                p['owner_name'] = data.get('owner_name', '') or p.get('owner_name', '')
                p['address'] = data.get('address', '') or p.get('address', '')
                p['company_name'] = data.get('company_name', '') or p.get('company_name', '')
                p['quotation_no'] = data.get('quotation_no', '') or p.get('quotation_no', '')
                p['quotation_date'] = data.get('date', '') or p.get('quotation_date', '')
                p['deposit'] = data.get('deposit', 0) or p.get('deposit', 0)
                # Recalculate total from items
                sub = sum((it.get('quantity', 1) or 1) * (it.get('unit_price', 0) or 0) for it in data['items'])
                p['total'] = sub
                p['quotation_total'] = sub
                p['updated_at'] = _now_iso()
                parsed += 1
        except Exception as e:
            # Mark as attempted to avoid retrying forever
            p['items'] = []
            p['_parse_error'] = str(e)[:200]
            print(f'Parse failed for {p.get("project_name", fid)}: {e}')
    if parsed > 0:
        _save(db)
    return parsed

def get_dashboard_stats():
    db = _load()
    projects = db['projects']
    now = datetime.now()
    this_month = now.strftime('%Y-%m')

    active = [p for p in projects if p.get('status') != 'cancelled']
    month_projects = [p for p in active if (p.get('quotation_date', '') or '')[:7] == this_month]

    total_quotation = sum(p.get('quotation_total', 0) for p in active)
    total_invoice = sum(p.get('invoice_total', 0) for p in active)
    total_deposit = sum(p.get('invoice_deposit', 0) for p in active)

    status_counts = {}
    for s in STATUSES:
        status_counts[s] = sum(1 for p in active if p.get('status') == s)

    cat_counts = {}
    cat_totals = {}
    for p in active:
        for it in p.get('items', []):
            cat = it.get('category', '雜項')
            cat_counts[cat] = cat_counts.get(cat, 0) + 1
            cat_totals[cat] = cat_totals.get(cat, 0) + (it.get('quantity', 1) or 1) * (it.get('unit_price', 0) or 0)

    top_cats = sorted(cat_counts.items(), key=lambda x: -x[1])[:5]

    monthly_revenue = {}
    for p in active:
        for dk in ['quotation_date', 'invoice_date']:
            d = (p.get(dk, '') or '')[:7]
            if d:
                monthly_revenue[d] = monthly_revenue.get(d, 0) + (p.get('quotation_total', 0) if dk == 'quotation_date' else p.get('invoice_total', 0))

    revenue_series = [{'month': k, 'total': v} for k, v in sorted(monthly_revenue.items())[-12:]]

    pending = 0
    for p in active:
        paid = p.get('payments_paid', [])
        pending += sum(1 for v in paid if not v)

    return {
        'total_projects': len(active),
        'this_month_count': len(month_projects),
        'this_month_total': sum(p.get('quotation_total', 0) for p in month_projects),
        'total_quotation_value': total_quotation,
        'total_invoice_value': total_invoice,
        'total_deposit_collected': total_deposit,
        'status_counts': status_counts,
        'top_categories': [{'name': c, 'count': n, 'total': cat_totals.get(c, 0)} for c, n in top_cats],
        'monthly_revenue': revenue_series,
        'pending_payments': pending,
        'confirmed_count': status_counts.get('confirmed', 0) + status_counts.get('in_progress', 0),
        'completed_count': status_counts.get('completed', 0),
    }
