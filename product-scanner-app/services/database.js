const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

let db;

function initDatabase() {
    const dbPath = path.join(__dirname, '..', 'data', 'mismatches.db');
    const dataDir = path.dirname(dbPath);
    
    // Create data directory if it doesn't exist
    if (!fs.existsSync(dataDir)) {
        fs.mkdirSync(dataDir, { recursive: true });
    }
    
    db = new sqlite3.Database(dbPath, (err) => {
        if (err) {
            console.error('Error opening database:', err);
        } else {
            console.log('Database connected');
            createTables();
        }
    });
}

function createTables() {
    db.run(`
        CREATE TABLE IF NOT EXISTS mismatches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            expected TEXT NOT NULL,
            actual TEXT NOT NULL,
            error_type TEXT DEFAULT 'mismatch',
            timestamp TEXT NOT NULL,
            status TEXT DEFAULT 'pending',
            resolved_at TEXT,
            resolved_by TEXT
        )
    `, (err) => {
        if (err) {
            console.error('Error creating table:', err);
        } else {
            addErrorTypeColumnIfMissing();
        }
    });
}

function addErrorTypeColumnIfMissing() {
    db.run(`ALTER TABLE mismatches ADD COLUMN error_type TEXT DEFAULT 'mismatch'`, (err) => {
        if (err && !err.message.includes('duplicate column')) {
            console.error('Error adding error_type column:', err);
        }
    });
}

function logMismatch(mismatchData) {
    return new Promise((resolve, reject) => {
        const { expected, actual, errorType = 'mismatch', timestamp, status = 'pending' } = mismatchData;
        
        db.run(
            `INSERT INTO mismatches (expected, actual, error_type, timestamp, status) 
             VALUES (?, ?, ?, ?, ?)`,
            [expected, actual, errorType, timestamp, status],
            function(err) {
                if (err) {
                    reject(err);
                } else {
                    resolve(this.lastID);
                }
            }
        );
    });
}

function getMismatches(limit = 500, startDate = null, endDate = null) {
    return new Promise((resolve, reject) => {
        let sql = `SELECT * FROM mismatches WHERE 1=1`;
        const params = [];

        if (startDate) {
            sql += ` AND date(timestamp) >= date(?)`;
            params.push(startDate);
        }
        if (endDate) {
            sql += ` AND date(timestamp) <= date(?)`;
            params.push(endDate);
        }
        sql += ` ORDER BY timestamp DESC LIMIT ?`;
        params.push(limit);

        db.all(sql, params, (err, rows) => {
            if (err) {
                reject(err);
            } else {
                resolve((rows || []).map(row => ({
                    id: row.id.toString(),
                    expected: row.expected,
                    actual: row.actual,
                    errorType: row.error_type || 'mismatch',
                    timestamp: row.timestamp,
                    status: row.status,
                    resolvedAt: row.resolved_at,
                    resolvedBy: row.resolved_by
                })));
            }
        });
    });
}

function getStatistics(startDate = null, endDate = null) {
    return new Promise((resolve, reject) => {
        let sql = `SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN date(timestamp) = date('now') THEN 1 ELSE 0 END) as today,
                SUM(CASE WHEN status IN ('sent', 'override') THEN 1 ELSE 0 END) as resolved,
                SUM(CASE WHEN status = 'pending' THEN 1 ELSE 0 END) as pending
             FROM mismatches`;
        const params = [];
        const conditions = [];
        if (startDate) {
            conditions.push('date(timestamp) >= date(?)');
            params.push(startDate);
        }
        if (endDate) {
            conditions.push('date(timestamp) <= date(?)');
            params.push(endDate);
        }
        if (conditions.length) sql += ' WHERE ' + conditions.join(' AND ');

        db.all(sql, params, (err, rows) => {
            if (err) {
                reject(err);
            } else {
                const stats = (rows && rows[0]) ? rows[0] : { total: 0, today: 0, resolved: 0, pending: 0 };
                resolve({
                    total: stats.total || 0,
                    today: stats.today || 0,
                    resolved: stats.resolved || 0,
                    pending: stats.pending || 0
                });
            }
        });
    });
}

function overrideMismatch(id) {
    return new Promise((resolve, reject) => {
        db.run(
            `UPDATE mismatches 
             SET status = 'override', 
                 resolved_at = datetime('now'),
                 resolved_by = 'operator'
             WHERE id = ?`,
            [id],
            function(err) {
                if (err) {
                    reject(err);
                } else {
                    resolve(this.changes);
                }
            }
        );
    });
}

module.exports = {
    initDatabase,
    logMismatch,
    getMismatches,
    getStatistics,
    overrideMismatch
};
