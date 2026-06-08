"""
Database SQLite per il SaaS contabile.
Tabelle: utenti, aziende, movimenti, report
"""
import sqlite3, os, datetime, json

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "saas.db")

def get_conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS utenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            nome TEXT,
            password_hash TEXT NOT NULL,
            piano TEXT DEFAULT 'free',
            stripe_customer_id TEXT,
            stripe_subscription_id TEXT,
            creato_il TEXT DEFAULT (datetime('now')),
            attivo INTEGER DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS aziende (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            utente_id INTEGER NOT NULL,
            nome TEXT NOT NULL,
            partita_iva TEXT,
            tipo TEXT DEFAULT 'ditta_individuale',
            creata_il TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(utente_id) REFERENCES utenti(id)
        );
        CREATE TABLE IF NOT EXISTS movimenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            azienda_id INTEGER NOT NULL,
            data TEXT,
            descrizione TEXT,
            importo REAL,
            imponibile REAL,
            aliquota_iva REAL DEFAULT 22.0,
            iva_importo REAL,
            categoria TEXT,
            commessa TEXT,
            tipo_documento TEXT,
            numero_documento TEXT,
            scadenza TEXT,
            stato_pagamento TEXT,
            fonte TEXT,
            creato_il TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(azienda_id) REFERENCES aziende(id)
        );
        CREATE TABLE IF NOT EXISTS report_generati (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            azienda_id INTEGER NOT NULL,
            utente_id INTEGER NOT NULL,
            tipo TEXT,
            periodo_da TEXT,
            periodo_a TEXT,
            path_xlsx TEXT,
            path_pdf TEXT,
            creato_il TEXT DEFAULT (datetime('now')),
            FOREIGN KEY(azienda_id) REFERENCES aziende(id)
        );
    """)
    conn.commit(); conn.close()

# ── UTENTI ──────────────────────────────────────────────
def crea_utente(email, nome, password_hash):
    conn = get_conn()
    try:
        conn.execute("INSERT INTO utenti(email,nome,password_hash) VALUES(?,?,?)",
                     (email.lower().strip(), nome, password_hash))
        conn.commit()
        return True, "OK"
    except sqlite3.IntegrityError:
        return False, "Email già registrata."
    finally:
        conn.close()

def get_utente_by_email(email):
    conn = get_conn()
    row = conn.execute("SELECT * FROM utenti WHERE email=? AND attivo=1",
                       (email.lower().strip(),)).fetchone()
    conn.close()
    return dict(row) if row else None

def aggiorna_piano(utente_id, piano, stripe_customer_id=None, stripe_sub_id=None):
    conn = get_conn()
    conn.execute("""UPDATE utenti SET piano=?,
                    stripe_customer_id=COALESCE(?,stripe_customer_id),
                    stripe_subscription_id=COALESCE(?,stripe_subscription_id)
                    WHERE id=?""",
                 (piano, stripe_customer_id, stripe_sub_id, utente_id))
    conn.commit(); conn.close()

# ── AZIENDE ─────────────────────────────────────────────
def crea_azienda(utente_id, nome, partita_iva="", tipo="ditta_individuale"):
    conn = get_conn()
    c = conn.execute("INSERT INTO aziende(utente_id,nome,partita_iva,tipo) VALUES(?,?,?,?)",
                     (utente_id, nome, partita_iva, tipo))
    conn.commit(); aid = c.lastrowid; conn.close()
    return aid

def get_aziende(utente_id):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM aziende WHERE utente_id=? ORDER BY nome",
                        (utente_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_azienda(azienda_id, utente_id):
    conn = get_conn()
    row = conn.execute("SELECT * FROM aziende WHERE id=? AND utente_id=?",
                       (azienda_id, utente_id)).fetchone()
    conn.close()
    return dict(row) if row else None

def elimina_azienda(azienda_id, utente_id):
    conn = get_conn()
    conn.execute("DELETE FROM movimenti WHERE azienda_id=?", (azienda_id,))
    conn.execute("DELETE FROM aziende WHERE id=? AND utente_id=?", (azienda_id, utente_id))
    conn.commit(); conn.close()

# ── MOVIMENTI ────────────────────────────────────────────
def salva_movimenti(azienda_id, movimenti):
    conn = get_conn()
    for m in movimenti:
        conn.execute("""INSERT INTO movimenti
            (azienda_id,data,descrizione,importo,imponibile,aliquota_iva,iva_importo,
             categoria,commessa,tipo_documento,numero_documento,scadenza,stato_pagamento,fonte)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (azienda_id, m.get("data",""), m.get("descrizione",""),
             m.get("importo",0), m.get("imponibile",0), m.get("aliquota_iva",22),
             m.get("iva_importo",0), m.get("categoria",""), m.get("commessa",""),
             m.get("tipo_documento",""), m.get("numero_documento",""),
             m.get("scadenza",""), m.get("stato_pagamento",""), m.get("fonte","")))
    conn.commit(); conn.close()

def get_movimenti(azienda_id, da=None, a=None, commessa=None):
    conn = get_conn()
    q = "SELECT * FROM movimenti WHERE azienda_id=?"
    params = [azienda_id]
    if da: q += " AND data >= ?"; params.append(da)
    if a: q += " AND data <= ?"; params.append(a)
    if commessa: q += " AND LOWER(commessa)=?"; params.append(commessa.lower())
    q += " ORDER BY data"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def cancella_movimenti(azienda_id):
    conn = get_conn()
    conn.execute("DELETE FROM movimenti WHERE azienda_id=?", (azienda_id,))
    conn.commit(); conn.close()

def salva_report(azienda_id, utente_id, tipo, da, a, path_xlsx, path_pdf):
    conn = get_conn()
    conn.execute("""INSERT INTO report_generati
        (azienda_id,utente_id,tipo,periodo_da,periodo_a,path_xlsx,path_pdf)
        VALUES(?,?,?,?,?,?,?)""",
        (azienda_id, utente_id, tipo, da, a, path_xlsx, path_pdf))
    conn.commit(); conn.close()

def get_report(azienda_id):
    conn = get_conn()
    rows = conn.execute("""SELECT * FROM report_generati WHERE azienda_id=?
                           ORDER BY creato_il DESC LIMIT 20""",
                        (azienda_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

if __name__ == "__main__":
    init_db(); print("DB inizializzato.")
