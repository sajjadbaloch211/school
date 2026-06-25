console.log({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  database: process.env.DB_NAME,
  port: process.env.DB_PORT,
  pass: process.env.DB_PASS ? "FOUND" : "MISSING"
});
});
const express = require('express');
const mysql = require('mysql2');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const bodyParser = require('body-parser');
const dotenv = require('dotenv');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const nodemailer = require('nodemailer');
const geoip = require('geoip-lite');
const UAParser = require('ua-parser-js');
const requestIp = require('request-ip');
const crypto = require('crypto');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const ExcelJS = require('exceljs'); // 📊 For Excel Export
const csvParser = require('csv-parser'); // 📥 For CSV Import

// 🔒 SECURITY PACKAGES
const csrf = require('csurf');
const cookieParser = require('cookie-parser');
const xss = require('xss');
const validator = require('validator');
const QRCode = require('qrcode');

// 🔒 ENTERPRISE SECURITY VAULT
const vault = require('./security_vault');


const auditSystem = require('./audit_system');

dotenv.config();

const app = express();
app.set('trust proxy', 1);

// Database Connection (MySQL/XAMPP)
const db = mysql.createPool({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'root',
    password: process.env.DB_PASS || '',
    database: process.env.DB_NAME || 'waqar_school_db',
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

// 🎨 Make db globally accessible for audit system
global.db = db;

// Verify Connection & Init
db.getConnection((err, connection) => {
    if (err) {
        console.error('Database connection failed:', err);
    } else {
        connection.release(); // Release immediately, just checking connectivity
        console.log('Connected to MySQL Database via XAMPP (Pool Mode)');

        // 🏢 MULTI-CAMPUS ARCHITECTURE
        const initQueries = [
            // 1. CAMPUSES TABLE
            `CREATE TABLE IF NOT EXISTS campuses (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_name VARCHAR(100) NOT NULL,
                campus_code VARCHAR(20) UNIQUE NOT NULL,
                address VARCHAR(255),
                city VARCHAR(100),
                contact_phone VARCHAR(20),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Library Table (Updated with campus_id)
            `CREATE TABLE IF NOT EXISTS library (
                id int(11) NOT NULL AUTO_INCREMENT,
                campus_id INT DEFAULT 1, 
                title varchar(255) NOT NULL,
                subject varchar(100),
                file_path varchar(255) NOT NULL,
                file_type varchar(20),
                uploaded_by int(11) NOT NULL,
                uploaded_at timestamp DEFAULT CURRENT_TIMESTAMP,
                target_grade VARCHAR(50) DEFAULT 'General',
                PRIMARY KEY (id),
                FOREIGN KEY (uploaded_by) REFERENCES users(id),
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE SET NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Timetable Table
            `CREATE TABLE IF NOT EXISTS timetable (
                id int(11) NOT NULL AUTO_INCREMENT,
                campus_id INT DEFAULT 1,
                class_id int(11) NOT NULL,
                subject varchar(100) NOT NULL,
                day varchar(20) NOT NULL,
                start_time varchar(20) NOT NULL,
                end_time varchar(20) NOT NULL,
                teacher_id int(11) DEFAULT NULL,
                created_at timestamp DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (id),
                FOREIGN KEY (class_id) REFERENCES classes(id) ON DELETE CASCADE,
                FOREIGN KEY (teacher_id) REFERENCES teachers(id) ON DELETE SET NULL,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Exams Table
            `CREATE TABLE IF NOT EXISTS exams (
                id int(11) NOT NULL AUTO_INCREMENT,
                campus_id INT DEFAULT 1,
                exam_name varchar(100) NOT NULL,
                start_date date NOT NULL,
                end_date date,
                created_at timestamp DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (id),
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Exam Results Table
            `CREATE TABLE IF NOT EXISTS exam_results (
                id int(11) NOT NULL AUTO_INCREMENT,
                campus_id INT DEFAULT 1,
                exam_id int(11) NOT NULL,
                student_id int(11) NOT NULL,
                subject varchar(100) NOT NULL,
                marks_obtained decimal(5,2) NOT NULL,
                total_marks int(11) NOT NULL,
                grade varchar(5),
                created_at timestamp DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (id),
                FOREIGN KEY (exam_id) REFERENCES exams(id) ON DELETE CASCADE,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Attendance Table
            `CREATE TABLE IF NOT EXISTS attendance (
                id int(11) NOT NULL AUTO_INCREMENT,
                campus_id INT DEFAULT 1,
                student_id int(11) NOT NULL,
                date date NOT NULL,
                status enum('present','absent','late','leave') DEFAULT 'present',
                marked_by int(11) NOT NULL,
                created_at timestamp DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (id),
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (marked_by) REFERENCES users(id),
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Homework Table
            `CREATE TABLE IF NOT EXISTS homework (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                teacher_id INT NOT NULL,
                class_id INT NOT NULL,
                subject VARCHAR(100) NOT NULL,
                title VARCHAR(255) NOT NULL,
                description TEXT,
                due_date DATE NOT NULL,
                total_marks INT DEFAULT 10,
                attachment_path VARCHAR(255),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (teacher_id) REFERENCES teachers(id) ON DELETE CASCADE,
                FOREIGN KEY (class_id) REFERENCES classes(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Homework Submissions Table
            `CREATE TABLE IF NOT EXISTS homework_submissions (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                homework_id INT NOT NULL,
                student_id INT NOT NULL,
                submission_file VARCHAR(255) NOT NULL,
                submitted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status ENUM('on_time', 'late') DEFAULT 'on_time',
                marks_obtained DECIMAL(5,2),
                teacher_remarks TEXT,
                graded_at DATETIME,
                FOREIGN KEY (homework_id) REFERENCES homework(id) ON DELETE CASCADE,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Payroll Table
            `CREATE TABLE IF NOT EXISTS payroll (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                teacher_id INT,
                amount DECIMAL(10,2),
                month VARCHAR(20),
                year INT,
                status ENUM('paid', 'pending') DEFAULT 'pending',
                paid_date DATE,
                FOREIGN KEY (teacher_id) REFERENCES teachers(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Fees Table (Monthly Line Items)
            `CREATE TABLE IF NOT EXISTS fees (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                voucher_id INT,
                student_id INT,
                amount DECIMAL(10,2) DEFAULT 0,
                month VARCHAR(20),
                year INT,
                status ENUM('paid', 'unpaid', 'partially_paid') DEFAULT 'unpaid',
                due_date DATE,
                carried_due DECIMAL(10,2) DEFAULT 0,
                line_total DECIMAL(10,2) DEFAULT 0,
                previous_dues DECIMAL(10,2) DEFAULT 0,
                total_amount DECIMAL(10,2) DEFAULT 0,
                paid_amount DECIMAL(10,2) DEFAULT 0,
                remaining_amount DECIMAL(10,2) DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Fee Payments Table (Transaction History)
            `CREATE TABLE IF NOT EXISTS fee_payments (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                fee_id INT,
                voucher_id INT,
                amount_paid DECIMAL(10,2) NOT NULL,
                payment_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                payment_method VARCHAR(50) DEFAULT 'Cash',
                receipt_no VARCHAR(50),
                recorded_by INT,
                FOREIGN KEY (fee_id) REFERENCES fees(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // MASTER VOUCHER TABLE (Ledger) - ONE PER YEAR
            `CREATE TABLE IF NOT EXISTS vouchers (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                student_id INT NOT NULL,
                academic_year INT NOT NULL,
                total_charged DECIMAL(10,2) DEFAULT 0,
                total_paid DECIMAL(10,2) DEFAULT 0,
                remaining_balance DECIMAL(10,2) DEFAULT 0,
                status ENUM('open', 'closed') DEFAULT 'open',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(student_id, academic_year),
                FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // Notices Table
            `CREATE TABLE IF NOT EXISTS notices (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                title VARCHAR(255) NOT NULL,
                content TEXT NOT NULL,
                posted_by INT,
                attachment_path VARCHAR(255),
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (posted_by) REFERENCES users(id) ON DELETE SET NULL,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`,

            // SECURITY AUDIT LOGS TABLE - 🔒 TAMPER PROOF
            `CREATE TABLE IF NOT EXISTS audit_logs (
                id INT AUTO_INCREMENT PRIMARY KEY,
                campus_id INT DEFAULT 1,
                user_id INT,
                role VARCHAR(20),
                action VARCHAR(50),
                ip_address VARCHAR(45),
                latitude DECIMAL(10, 8),
                longitude DECIMAL(11, 8),
                country VARCHAR(100),
                city VARCHAR(100),
                device VARCHAR(100),
                browser VARCHAR(100),
                os VARCHAR(100),
                environment VARCHAR(10),
                distance_from_last_login INT DEFAULT 0,
                risk_level ENUM('Low', 'Medium', 'High') DEFAULT 'Low',
                integrity_hash VARCHAR(64), -- 🔒 Hash of (data + secret)
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL,
                FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE SET NULL
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;`
        ];

        // Execute all init queries
        initQueries.forEach(query => {
            db.query(query, (err) => {
                if (err) console.error("Database Init Error:", err.message);
            });
        });

        // 🔗 MULTI-CAMPUS MIGRATION & KEYS
        setTimeout(() => {
            // 1. Create Default Campus
            db.query(`INSERT IGNORE INTO campuses (id, campus_name, campus_code, address) VALUES (1, 'Main Campus', 'MAIN', 'HQ Address')`, (err) => { });

            // 2. Add campus_id to Users Table (Critical for Login)
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS campus_id INT DEFAULT 1", (err) => {
                db.query("ALTER TABLE users ADD CONSTRAINT fk_user_campus FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE SET NULL", (e) => { });

                // 🔐 RULE 5: Allow Same User ID in different campuses
                // We attempt to drop the global 'username' unique index and replace it with (username + campus_id)
                db.query("SHOW INDEX FROM users WHERE Key_name = 'username'", (errShow, indexResults) => {
                    if (indexResults && indexResults.length > 0) {
                        db.query("ALTER TABLE users DROP INDEX username", (errDrop) => {
                            if (errDrop) console.log("Note: Could not drop username index (might be FK or already dropped).");

                            // Add new composite unique index
                            db.query("ALTER TABLE users ADD UNIQUE INDEX unique_user_campus (username, campus_id)", (errAdd) => {
                                if (errAdd && errAdd.code !== 'ER_DUP_KEYNAME') console.log("Note: User uniqueness constraint updated.");
                            });
                        });
                    } else {
                        // Index already dropped or doesn't exist, just try adding new one
                        db.query("ALTER TABLE users ADD UNIQUE INDEX unique_user_campus (username, campus_id)", (errAdd) => { });
                    }
                });
            });

            // 3. Add campus_id to other Core Entities
            const tablesWithCampus = ['students', 'teachers', 'classes', 'timetable'];
            tablesWithCampus.forEach(table => {
                db.query(`ALTER TABLE ${table} ADD COLUMN IF NOT EXISTS campus_id INT DEFAULT 1`, (err) => {
                    db.query(`ALTER TABLE ${table} ADD CONSTRAINT fk_${table}_campus FOREIGN KEY (campus_id) REFERENCES campuses(id) ON DELETE CASCADE`, (e) => { });
                });
            });

            // 🔐 BANK-GRADE SECURITY UPGRADES
            // 1. MFA Support in Users Table
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS mfa_secret VARCHAR(64) DEFAULT NULL", (err) => { });
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS mfa_enabled TINYINT(1) DEFAULT 0", (err) => { });
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS last_password_change TIMESTAMP DEFAULT CURRENT_TIMESTAMP", (err) => { });
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS force_password_reset TINYINT(1) DEFAULT 0", (err) => { });

            // 2. Account Lockout Support
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS login_attempts INT DEFAULT 0", (err) => { });
            db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS lockout_until TIMESTAMP NULL DEFAULT NULL", (err) => { });

        }, 2000);

        // 🟢 STRICT LEDGER CONSTRAINTS (SQL Level Protection)
        db.query("ALTER TABLE vouchers ADD UNIQUE IF NOT EXISTS unique_student_year (student_id, academic_year)", (err) => { });
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS voucher_id INT", (err) => {
            // Unique month per voucher protection
            db.query("ALTER TABLE fees ADD UNIQUE IF NOT EXISTS unique_voucher_month (voucher_id, month, year)", (err2) => { });
        });
        db.query("ALTER TABLE fee_payments ADD COLUMN IF NOT EXISTS voucher_id INT", (err) => { });

        // Add Ledger Columns to Fees (Line Items)
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS carried_due DECIMAL(10,2) DEFAULT 0", (err) => { });
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS line_total DECIMAL(10,2) DEFAULT 0", (err) => { });

        // Add salary column to teachers if it doesn't exist
        db.query("ALTER TABLE teachers ADD COLUMN IF NOT EXISTS salary DECIMAL(10,2) DEFAULT 0", (err) => {
            if (err) console.log("Note: Salary column check.");
        });

        // Add visibility column to library if it doesn't exist
        db.query("ALTER TABLE library ADD COLUMN IF NOT EXISTS visibility ENUM('student', 'teacher', 'both') DEFAULT 'both'", (err) => {
            if (err) console.log("Note: Visibility column check.");
        });

        // Add monthly_fee column to classes if it doesn't exist
        db.query("ALTER TABLE classes ADD COLUMN IF NOT EXISTS monthly_fee DECIMAL(10,2) DEFAULT 0", (err) => {
            if (err) console.log("Note: Classes monthly_fee column check.");
        });

        // Add attachment_path column to notices safely
        db.query("ALTER TABLE notices ADD COLUMN attachment_path VARCHAR(255)", (err) => {
            if (err && err.code !== 'ER_DUP_FIELDNAME') console.log("Note: " + err.message);
        });

        // 🛠️ LEDGER MIGRATION: Auto-link orphan records to Master Vouchers
        setTimeout(() => {
            console.log("Running Ledger Migration...");
            // 1. Create missing vouchers for any student who has fee entries
            const createMissingVouchers = `
                INSERT IGNORE INTO vouchers (student_id, academic_year, total_charged, total_paid, remaining_balance)
                SELECT DISTINCT student_id, year, 0, 0, 0 FROM fees WHERE voucher_id IS NULL
            `;
            db.query(createMissingVouchers, () => {
                // 2. Link fees to vouchers
                const linkFees = `
                    UPDATE fees f
                    JOIN vouchers v ON f.student_id = v.student_id AND f.year = v.academic_year
                    SET f.voucher_id = v.id
                    WHERE f.voucher_id IS NULL
                `;
                db.query(linkFees, () => {
                    // 3. Link payments to vouchers
                    const linkPayments = `
                        UPDATE fee_payments fp
                        JOIN fees f ON fp.fee_id = f.id
                        SET fp.voucher_id = f.voucher_id
                        WHERE fp.voucher_id IS NULL
                    `;
                    db.query(linkPayments, () => {
                        // 4. Force global recalculation
                        db.query("SELECT id FROM vouchers", (err, rows) => {
                            if (rows) rows.forEach(v => recalculateVoucher(v.id));
                            console.log("Ledger Migration Complete.");
                        });
                    });
                });
            });
        }, 3000);


        // Add partial payment columns to fees
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS previous_dues DECIMAL(10,2) DEFAULT 0", (err) => { });
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS total_amount DECIMAL(10,2) DEFAULT 0", (err) => { });
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS paid_amount DECIMAL(10,2) DEFAULT 0", (err) => { });
        db.query("ALTER TABLE fees ADD COLUMN IF NOT EXISTS remaining_amount DECIMAL(10,2) DEFAULT 0", (err) => { });

        // Update Status ENUM to include partially_paid
        db.query("ALTER TABLE fees MODIFY COLUMN status ENUM('paid', 'unpaid', 'partially_paid') DEFAULT 'unpaid'", (err) => {
            if (err) console.log("Note: Fees status column update.");
        });

        // DATA MIGRATION: Initialize remaining_amount for old records if they are zero but unpaid
        db.query("UPDATE fees SET remaining_amount = total_amount WHERE (remaining_amount = 0 OR remaining_amount IS NULL) AND status = 'unpaid' AND total_amount > 0", (err) => {
            if (err) console.error("Migration Error:", err);
        });

        // Add Unique Constraint to Fees to prevent duplicates [student_id, month, year]
        db.query("ALTER TABLE fees ADD UNIQUE INDEX student_month_year (student_id, month, year)", (err) => {
            if (err) {
                // Ignore error if it already exists
                if (err.code !== 'ER_DUP_KEYNAME') console.log("Note: Fees unique index check.");
            }
        });

        // AUTH UPDATE: Add Google ID and Reset Token columns
        db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS google_id VARCHAR(255) DEFAULT NULL", (err) => { });
        db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS reset_token VARCHAR(255) DEFAULT NULL", (err) => { });
        db.query("ALTER TABLE users ADD COLUMN IF NOT EXISTS reset_token_expires DATETIME DEFAULT NULL", (err) => { });

        console.log("Database tables verified.");
    }
});



// Global Website Data Middleware (Theme, Settings, Menus)
// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));

// 🔒 Security Headers (Bank-Grade Hardware-Accelerated)
app.use(helmet({
    contentSecurityPolicy: false, // EJS Compatibility
    crossOriginEmbedderPolicy: false,
    strictTransportSecurity: {
        maxAge: 31536000,
        includeSubDomains: true,
        preload: true
    },
    xFrameOptions: { action: 'deny' },
    xContentTypeOptions: true,
    referrerPolicy: { policy: 'same-origin' }
}));

// 🔒 Cookie Parser (Required for CSRF)
app.use(cookieParser());

// 🔒 Rate Limiters
const authLimiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 5, // ✅ SECURITY FIX: Reduced from 15 to 5 attempts
    message: 'Too many login attempts from this IP, please try again after 15 minutes',
    standardHeaders: true,
    legacyHeaders: false,
    skipSuccessfulRequests: true // Don't count successful logins
});

// 🔒 Rate Limiter for Data Modification
const dataModifyLimiter = rateLimit({
    windowMs: 60 * 1000, // 1 minute
    max: 30, // 30 requests per minute
    message: 'Too many requests, please slow down'
});

// 🔒 Rate Limiter for Exports
const exportLimiter = rateLimit({
    windowMs: 60 * 1000, // 1 minute
    max: 5, // 5 exports per minute
    message: 'Too many export requests'
});

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// 🔒 Session Configuration
app.use(session({
    secret: process.env.SESSION_SECRET || 'waqar_secret_key',
    resave: false,
    saveUninitialized: false,
    cookie: {
        maxAge: 3600000, // 1 hour
        httpOnly: true, // Prevents client-side JS from reading the cookie
        secure: false,// ✅ SECURITY FIX: Enable in production
        sameSite: 'lax' // ✅ SECURITY FIX: CSRF protection
    }
}));

// 🔒 CSRF Protection
const csrfProtection = csrf({ cookie: true });

// Apply Rate Limiting to Auth Routes
app.use('/auth/', authLimiter);






// 🎨 Load CMS Globals (Settings, Theme, Menus)
app.use((req, res, next) => {
    res.locals.user = req.session.user || null;
    res.locals.campus = req.session.campus || null;
    res.locals.csrfToken = null;

    // 🔒 ZERO-TRUST: Device Fingerprint Check
    if (req.session.user) {
        const currentFingerprint = vault.generateDeviceFingerprint(req);
        if (req.session.fingerprint && req.session.fingerprint !== currentFingerprint) {
            console.warn(`[ZERO-TRUST] Fingerprint Mismatch for UID: ${req.session.user.id}. Session Destroyed.`);
            return req.session.destroy(() => res.redirect('/login?error=Session+Security+Violation'));
        }
    }

    // 1. Get Site Settings
    db.query('SELECT setting_key, setting_value FROM cms_settings', (err, settingResults) => {
        const settings = {};
        if (!err && settingResults) {
            settingResults.forEach(r => settings[r.setting_key] = r.setting_value);
        }
        res.locals.settings = settings;

        // 2. Get Active Theme
        db.query('SELECT colors, fonts, button_styles FROM cms_themes WHERE is_active = TRUE LIMIT 1', (err, themeResults) => {
            let theme = {
                colors: { primary: '#1e3a8a', secondary: '#3b82f6', accent: '#60a5fa' },
                fonts: { heading: 'Inter, sans-serif', body: 'Inter, sans-serif' },
                button_styles: { borderRadius: '8px', shadow: 'none' }
            };
            if (!err && themeResults && themeResults.length > 0) {
                try {
                    const row = themeResults[0];
                    theme.colors = typeof row.colors === 'string' ? JSON.parse(row.colors) : row.colors;
                    theme.fonts = typeof row.fonts === 'string' ? JSON.parse(row.fonts) : row.fonts;
                    theme.button_styles = typeof row.button_styles === 'string' ? JSON.parse(row.button_styles) : (row.button_styles || theme.button_styles);
                } catch (e) {
                    console.error("Theme parse error:", e);
                }
            }
            res.locals.theme = theme;

            // 3. Get Menus
            db.query('SELECT label, url FROM cms_menu_items WHERE menu_location = "header" AND enabled = TRUE ORDER BY display_order', (err, hMenu) => {
                res.locals.headerMenu = err ? [] : hMenu;
                db.query('SELECT label, url FROM cms_menu_items WHERE menu_location = "footer" AND enabled = TRUE ORDER BY display_order', (err, fMenu) => {
                    res.locals.footerMenu = err ? [] : fMenu;

                    // 4. Get Current Page Title (helper for views)
                    res.locals.page = { title: settings.site_name || 'Waqar Public Higher Secondary School' };
                    next();
                });
            });
        });
    });
});

// 🔒 GLOBAL CSRF PROTECTION & TOKEN GENERATION
app.use((req, res, next) => {
    // Skip for static assets
    if (req.path.includes('.') || req.path.startsWith('/uploads')) {
        return next();
    }

    if (req.method === 'GET') {
        csrfProtection(req, res, (err) => {
            if (!err && typeof req.csrfToken === 'function') {
                res.locals.csrfToken = req.csrfToken();
            }
            next(); // Proceed regardless of error on GET
        });
    } else {
        // POST/PUT/DELETE: CSRF protection is applied via route-specific middleware or here
        // We already added it to most routes, but we'll set the local for re-renders on error
        next();
    }
});



// 🔒 Multer Setup for Library with Security
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const dir = './public/uploads/library';
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        cb(null, dir);
    },
    filename: (req, file, cb) => {
        const sanitizedName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, '_');
        cb(null, Date.now() + '-' + sanitizedName);
    }
});

const upload = multer({
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 },
    fileFilter: (req, file, cb) => {
        const allowedMimes = ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png', 'image/gif', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'video/mp4', 'video/avi', 'video/quicktime', 'application/zip', 'application/x-rar-compressed'];
        if (allowedMimes.includes(file.mimetype)) cb(null, true);
        else cb(new Error(`Invalid file type: ${file.mimetype}`));
    }
});

// 📁 Homework Storage Setup
const homeworkStorage = multer.diskStorage({
    destination: (req, file, cb) => {
        const dir = './public/uploads/homework';
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        cb(null, dir);
    },
    filename: (req, file, cb) => {
        const sanitizedName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, '_');
        cb(null, Date.now() + '-hw-' + sanitizedName);
    }
});
const uploadHomework = multer({ storage: homeworkStorage, limits: { fileSize: 10 * 1024 * 1024 } });

// 📁 Student Submission Storage Setup
const submissionStorage = multer.diskStorage({
    destination: (req, file, cb) => {
        const dir = './public/uploads/submissions';
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        cb(null, dir);
    },
    filename: (req, file, cb) => {
        const sanitizedName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, '_');
        cb(null, Date.now() + '-sub-' + sanitizedName);
    }
});
const uploadSubmission = multer({ storage: submissionStorage, limits: { fileSize: 10 * 1024 * 1024 } });

// Routes (MIGRATE TO DYNAMIC CMS)
app.get('/', (req, res) => res.render('index', { isHomepage: true }));
app.get('/about', (req, res) => res.render('about'));
app.get('/admissions', (req, res) => res.render('admissions'));
app.get('/fee-structure', (req, res) => res.render('fee-structure'));
app.get('/gallery', (req, res) => res.render('gallery'));
app.get('/contact', (req, res) => res.render('contact'));

// Login Page
app.get('/login', csrfProtection, (req, res) => {
    if (req.session.user) {
        const returnTo = req.query.return_to || `/${req.session.user.role}/dashboard`;
        return res.redirect(returnTo);
    }

    const campusCode = req.query.campus_code || 'MAIN';
    const returnTo = req.query.return_to || null;

    // 🛡️ Handle Redirect Errors
    let errorMsg = null;
    if (req.query.error === 'session_expired') {
        errorMsg = 'Security session expired. Please try again.';
    } else if (req.query.error) {
        errorMsg = req.query.error;
    }

    const protocol = req.protocol;
    const host = req.get('host');
    const loginUrl = `${protocol}://${host}/login?campus_code=${campusCode}`;

    QRCode.toDataURL(loginUrl, (err, qrCodeUrl) => {
        logSecurityEvent(req, null, 'VIEW_LOGIN');
        res.render('login', {
            error: errorMsg,
            qrCode: qrCodeUrl || null,
            campusCode: campusCode,
            returnTo: returnTo,
            csrfToken: req.csrfToken()
        });
    });
});

// Login Logic
// 🔒 CSRF Protected
app.post('/auth/login', csrfProtection, (req, res) => {
    let { username, password, campus_code } = req.body;

    // Default to 'MAIN' if not provided
    campus_code = campus_code ? campus_code.trim().toUpperCase() : 'MAIN';
    username = username ? username.trim() : '';

    if (!username || !password) {
        logSecurityEvent(req, null, 'FAILED_LOGIN_MISSING_CREDENTIALS');
        return res.render('login', { error: 'Username and Password are required.' });
    }

    // 1. Verify Campus First
    const campusQuery = "SELECT * FROM campuses WHERE campus_code = ?";
    db.query(campusQuery, [campus_code], (err, campuses) => {
        if (err || campuses.length === 0) {
            logSecurityEvent(req, null, 'FAILED_LOGIN_INVALID_CAMPUS');
            return res.render('login', { error: 'Invalid Campus Code' });
        }

        const currentCampus = campuses[0];

        // 2. Find User WITHIN this Campus
        let query = '';
        let params = [];

        if (username.includes('@')) {
            // Email Login (Scoped to Campus)
            query = `SELECT * FROM users WHERE email = ? AND campus_id = ?`;
            params = [username, currentCampus.id];
        } else {
            // Username Login (Scoped to Campus)
            query = 'SELECT * FROM users WHERE username = ? AND campus_id = ?';
            params = [username, currentCampus.id];
        }

        db.query(query, params, async (err, results) => {
            if (err) {
                logSecurityEvent(req, null, 'SYSTEM_ERROR_LOGIN');
                console.error(err);
                return res.render('login', { error: 'Internal Server Error' });
            }

            if (results.length === 0) {
                logSecurityEvent(req, null, 'FAILED_LOGIN_USER_NOT_FOUND_IN_CAMPUS');
                return res.render('login', { error: 'Invalid Credentials' });
            }

            const user = results[0];

            // 1. 🔒 ACCOUNT LOCKOUT CHECK
            if (user.lockout_until && new Date() < new Date(user.lockout_until)) {
                logSecurityEvent(req, user, 'LOGIN_BLOCKED_LOCKOUT');
                return res.render('login', { error: 'Account locked due to multiple failed attempts. Try later.' });
            }

            // 2. 🔒 PASSWORD VERIFICATION
            const match = await bcrypt.compare(password, user.password);

            if (match) {
                // 🔒 Reset failed attempts on success
                db.query("UPDATE users SET login_attempts = 0, lockout_until = NULL WHERE id = ?", [user.id]);

                // 🔒 Check for MFA (2FA)
                if (user.mfa_enabled) {
                    req.session.mfa_user_id = user.id; // Temporary session for MFA
                    return res.redirect('/verify-mfa');
                }

                // 🔒 LOGIN SUCCESS - COMPLETE SESSION HARDENING
                req.session.regenerate((err) => { // 🔒 Session Rotation
                    if (err) return res.render('login', { error: 'Session Error' });

                    req.session.user = user;
                    req.session.campus = currentCampus;

                    // 🔒 DEVICE BINDING
                    req.session.fingerprint = vault.generateDeviceFingerprint(req);

                    logSecurityEvent(req, user, 'LOGIN_SUCCESS');
                    const returnTo = req.body.return_to || `/${user.role}/dashboard`;
                    return res.redirect(returnTo);
                });

            } else {
                // 🔒 FAILED ATTEMPT - Increment Counter
                const attempts = (user.login_attempts || 0) + 1;
                let lockoutReset = null;

                if (attempts >= 5) { // 🔒 Lock after 5 attempts
                    const lockDate = new Date();
                    lockDate.setMinutes(lockDate.getMinutes() + 15); // 15 min lock
                    lockoutReset = lockDate;
                    logSecurityEvent(req, user, 'ACCOUNT_LOCKED');
                }

                db.query("UPDATE users SET login_attempts = ?, lockout_until = ? WHERE id = ?", [attempts, lockoutReset, user.id]);

                logSecurityEvent(req, user, 'FAILED_LOGIN_BAD_PASSWORD');
                return res.render('login', { error: 'Invalid Credentials' });
            }
        });
    });
});

// 🔒 MFA Verification Page (Post-Login)
app.get('/verify-mfa', (req, res) => {
    if (!req.session.mfa_user_id) return res.redirect('/login');

    const protocol = req.protocol;
    const host = req.get('host');
    const qrSessionToken = crypto.randomBytes(32).toString('hex');
    const authUrl = `${protocol}://${host}/auth/qr-authorize?token=${qrSessionToken}`;

    // Store in DB with 5-minute expiry
    const expiry = new Date(Date.now() + 5 * 60 * 1000);
    db.execute(
        'INSERT INTO qr_sessions (token, expires_at) VALUES (?, ?)',
        [qrSessionToken, expiry],
        (err) => {
            if (err) console.error("[QR AUTH ERROR]", err);

            QRCode.toDataURL(authUrl, (qrErr, qrCodeUrl) => {
                res.render('mfa_verify', {
                    error: null,
                    qrCode: qrCodeUrl || null,
                    qrToken: qrSessionToken
                });
            });
        }
    );
});

// 📱 QR Authorization Route (Accessed via Phone)
app.get('/auth/qr-authorize', (req, res) => {
    if (!req.session.user) {
        // Not logged in on phone? Send to login but remember where to return
        return res.redirect(`/login?return_to=${encodeURIComponent(req.originalUrl)}`);
    }

    const { token } = req.query;
    const userId = req.session.user.id;

    console.log(`[QR AUTH] User ${userId} authorizing token ${token}`);

    db.execute(
        'UPDATE qr_sessions SET user_id = ?, status = "authorized" WHERE token = ? AND expires_at > NOW() AND status = "pending"',
        [userId, token],
        (err, result) => {
            if (err) {
                console.error(err);
                return res.send("System Error during authorization.");
            }

            if (result.affectedRows === 0) {
                return res.send("Invalid or expired QR code. Please refresh the desktop page.");
            }

            res.send(`
                <div style="font-family: sans-serif; text-align: center; padding: 50px;">
                    <div style="color: #10b981; font-size: 5rem; margin-bottom: 20px;"><i class="fas fa-check-circle"></i></div>
                    <h1 style="color: #0f172a;">Success!</h1>
                    <p style="color: #64748b;">The desktop session has been authorized. You can close this window now.</p>
                </div>
            `);
        }
    );
});

// 💻 QR Check Route (Polled by Desktop)
app.get('/auth/qr-check', (req, res) => {
    const { token } = req.query;
    if (!token) return res.json({ status: 'error' });

    db.execute(
        'SELECT u.*, q.status FROM qr_sessions q LEFT JOIN users u ON q.user_id = u.id WHERE q.token = ? AND q.expires_at > NOW()',
        [token],
        (err, results) => {
            if (err || results.length === 0) return res.json({ status: 'expired' });

            const qrSession = results[0];
            if (qrSession.status === 'authorized') {
                // Fulfill the session
                req.session.user = qrSession;
                delete req.session.user.password; // Security

                db.query("SELECT * FROM campuses WHERE id = ?", [qrSession.campus_id], (err2, campuses) => {
                    if (!err2 && campuses.length > 0) {
                        req.session.campus = campuses[0];
                        req.session.fingerprint = vault.generateDeviceFingerprint(req);
                        res.json({ status: 'authorized', role: qrSession.role });
                    } else {
                        res.json({ status: 'authorized', role: qrSession.role });
                    }
                });
            } else {
                res.json({ status: 'pending' });
            }
        }
    );
});

// 🔒 MFA Verification Logic
app.post('/auth/mfa-verify', csrfProtection, (req, res) => {
    if (!req.session.mfa_user_id) return res.redirect('/login');

    const { token } = req.body;
    const userId = req.session.mfa_user_id;

    db.query("SELECT * FROM users WHERE id = ?", [userId], (err, results) => {
        if (err || results.length === 0) return res.redirect('/login');
        const user = results[0];

        const verified = vault.verifyMFAToken(token, user.mfa_secret);

        if (verified) {
            // Success! Complete login
            db.query("SELECT * FROM campuses WHERE id = ?", [user.campus_id], (err2, campuses) => {
                req.session.regenerate((err3) => {
                    req.session.user = user;
                    req.session.campus = campuses[0];
                    req.session.fingerprint = vault.generateDeviceFingerprint(req);
                    delete req.session.mfa_user_id; // Clear temp session

                    logSecurityEvent(req, user, 'LOGIN_SUCCESS_MFA');
                    return res.redirect(`/${user.role}/dashboard`);
                });
            });
        } else {
            logSecurityEvent(req, user, 'FAILED_MFA_CODE');
            res.render('mfa_verify', { error: 'Invalid 6-digit code. Please try again.', csrfToken: req.csrfToken() });
        }
    });
});

// 🔒 Admin MFA Setup Page
app.get('/admin/mfa-setup', csrfProtection, (req, res) => {
    console.log(`[DEBUG] MFA Setup Hit for User: ${req.session.user ? req.session.user.username : 'None'}`);
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    // Check if already enabled
    if (req.session.user.mfa_enabled || req.session.user.mfa_secret) {
        return res.redirect('/admin/dashboard');
    }

    // Generate Secret
    const mfa = vault.generateMFA(req.session.user.username);

    QRCode.toDataURL(mfa.otpauth_url, (err, data_url) => {
        res.render('admin/mfa_setup', {
            qrCode: data_url,
            secret: mfa.base32,
            csrfToken: req.csrfToken(),
            error: null
        });
    });
});

// 🔒 Admin MFA Enable Logic (First time verification)
app.post('/admin/mfa-enable', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const { token, secret } = req.body;
    const verified = vault.verifyMFAToken(token, secret);

    if (verified) {
        db.query("UPDATE users SET mfa_secret = ?, mfa_enabled = 1 WHERE id = ?",
            [secret, req.session.user.id], (err) => {
                if (err) return res.send("Error enabling MFA");
                req.session.user.mfa_enabled = 1; // Update session
                logSecurityEvent(req, req.session.user, 'MFA_ENABLED');
                res.redirect('/admin/dashboard?success=MFA+Enabled+Successfully');
            });
    } else {
        // Regeneration on failure
        const mfa = vault.generateMFA(req.session.user.username);
        QRCode.toDataURL(mfa.otpauth_url, (err, data_url) => {
            res.render('admin/mfa_setup', {
                error: 'Invalid code. Scan again and enter the current 6-digit code.',
                qrCode: data_url,
                secret: mfa.base32,
                csrfToken: req.csrfToken()
            });
        });
    }
});


// Forgot Password View
app.get('/forgot-password', (req, res) => {
    res.render('forgot-password');
});

// Forgot Password Logic (Email & Phone Simulation)
// =============================================================
// SECURE PASSWORD RESET FLOW (Strict 10-Step Protocol)
// =============================================================

// Step 1: User Request (ID Entry Only)
app.post('/auth/forgot-password', (req, res) => {
    const { username } = req.body; // Strictly accepts User ID / Student ID

    console.log(`[RESET START] Request for User ID: "${username}"`);

    // Step 2: Find User & Get UserID
    // Strict Query: Find user ONLY by Username (ID)
    const query = `
        SELECT u.id, u.email, u.full_name, s.phone as s_phone, t.phone as t_phone 
        FROM users u 
        LEFT JOIN students s ON u.id = s.user_id 
        LEFT JOIN teachers t ON u.id = t.user_id 
        WHERE u.username = ?
    `;

    db.query(query, [username], (err, results) => {
        if (err) {
            console.error("[RESET ERROR] DB Search Failed:", err);
            return res.render('forgot-password', { error: 'System Error. Please try again.' });
        }

        if (results.length === 0) {
            console.log(`[RESET FAILED] Invalid User ID: "${username}"`);
            // Explicit Error as requested: "Invalid User ID. Please check and try again."
            return res.render('forgot-password', { error: 'Invalid User ID. Please check and try again.' });
        }

        const user = results[0];
        console.log(`[RESET FOUND] UserID: ${user.id} | Name: ${user.full_name}`);

        // Step 3: Generate One-Time OTP
        const otp = Math.floor(100000 + Math.random() * 900000).toString();

        // Step 5: Set Expiry (10 Minutes from now)
        const expires = new Date(Date.now() + 10 * 60 * 1000);

        // Step 4 & 5: Save OTP + Expiry to Database (Bound to UserID)
        const updateQuery = 'UPDATE users SET reset_token = ?, reset_token_expires = ? WHERE id = ?';

        db.execute(updateQuery, [otp, expires, user.id], async (updateErr) => {
            if (updateErr) {
                console.error("[RESET ERROR] Token Save Failed:", updateErr);
                return res.render('forgot-password', { error: 'System Error during OTP generation.' });
            }

            // Step 6: Send OTP (Email)
            let emailSent = false;
            if (user.email) {
                const transporter = nodemailer.createTransport({
                    service: 'gmail',
                    auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS }
                });

                try {
                    await transporter.sendMail({
                        from: process.env.EMAIL_USER,
                        to: user.email,
                        subject: 'Password Reset Verification Code',
                        text: `Hello ${user.full_name},\n\n` +
                            `You have requested to reset your password.\n` +
                            `Use this One-Time Code (OTP) to verify your identity:\n\n` +
                            `Code: ${otp}\n\n` +
                            `This code matches User ID: ${user.id} and is valid for 10 minutes.\n` +
                            `If you did not request this, your account is safe, simply ignore this email.`
                    });
                    console.log(`[EMAIL SENT] To: ${user.email} | OTP: ${otp}`);
                    emailSent = true;
                } catch (emailErr) {
                    console.error("[RESET WARNING] Email Delivery Failed:", emailErr.message);
                }
            } else {
                console.log(`[RESET INFO] No email on file for UserID ${user.id}. Falling back to Console/Phone.`);
            }

            // Step 6 (Alternate): Console / Simulated SMS
            const verificationPhone = user.s_phone || user.t_phone;
            console.log("========================================");
            console.log(`>>> SECURITY ALERT: PASSWORD RESET OTP <<<`);
            console.log(`>>> User: ${user.full_name} (ID: ${user.id})`);
            console.log(`>>> OTP:  ${otp}`);
            if (verificationPhone) console.log(`>>> SMS To: ${verificationPhone}`);
            console.log("========================================");

            // Navigate to Verification Page
            res.render('verify-otp', {
                userId: user.id,
                error: null,
                success: emailSent ? 'Code sent to your email.' : 'Code sent to your phone (Simulated).'
            });
        });
    });
});

// Step 7 & 8: Verify OTP
app.post('/auth/verify-otp', (req, res) => {
    const { user_id, otp } = req.body;

    console.log(`[VERIFY START] Checking OTP for UserID: ${user_id}`);

    // Verify:
    // 1. OTP Matches
    // 2. OTP Not Expired (> NOW())
    // 3. UserID Matches
    const query = 'SELECT id FROM users WHERE id = ? AND reset_token = ? AND reset_token_expires > NOW()';

    db.query(query, [user_id, otp], (err, results) => {
        if (err) {
            console.error("[VERIFY ERROR] DB Error:", err);
            return res.render('verify-otp', { userId: user_id, error: 'System Verification Error' });
        }

        if (results.length === 0) {
            console.warn(`[VERIFY FAILED] Invalid or Expired OTP for UserID: ${user_id}`);
            return res.render('verify-otp', { userId: user_id, error: 'Invalid Code or Code Expired. Try Again.' });
        }

        console.log(`[VERIFY SUCCESS] Identity confirmed for UserID: ${user_id}`);

        // Render Change Password Screen (Passing valid token forward for final confirmation)
        res.render('reset-password', { token: otp, userId: user_id, error: null });
    });
});

// Step 9 & 10: Final Password Update
app.post('/auth/reset-password-final', async (req, res) => {
    const { userId, token, password, confirm_password } = req.body;

    console.log(`[UPDATE START] Setting new password for UserID: ${userId}`);

    // Standard Validation
    if (!password || password.length < 8) {
        return res.render('reset-password', { token, userId, error: 'Password must be at least 8 characters.' });
    }
    if (password !== confirm_password) {
        // user helper to keep them on page
        return res.render('reset-password', { token, userId, error: 'Passwords do not match.' });
    }

    // Double-Verify Token in DB (Prevent Hijacking)
    const checkQuery = 'SELECT id FROM users WHERE id = ? AND reset_token = ? AND reset_token_expires > NOW()';

    db.query(checkQuery, [userId, token], async (err, users) => {
        if (err || users.length === 0) {
            return res.redirect('/forgot-password?error=Session+Expired+or+Invalid+State');
        }

        try {
            // Hash new password
            const hashedPassword = await bcrypt.hash(password, 10);

            // ATOMIC UPDATE:
            // 1. Update Password
            // 2. Clear Token (Invalidate OTP)
            // 3. Clear Expiry
            const updateQuery = 'UPDATE users SET password = ?, reset_token = NULL, reset_token_expires = NULL WHERE id = ?';

            db.execute(updateQuery, [hashedPassword, userId], (upErr, result) => {
                if (upErr) {
                    console.error("[UPDATE ERROR] DB Write Failed:", upErr);
                    return res.render('reset-password', { token, userId, error: 'Database Write Failed.' });
                }

                if (result.affectedRows === 1) {
                    console.log(`[UPDATE SUCCESS] Password changed for UserID: ${userId}. OTP Invalidated.`);
                    res.redirect('/login?success=Password+Changed+Successfully.+Please+Login.');
                } else {
                    console.error("[UPDATE ERROR] No rows affected.");
                    res.render('reset-password', { token, userId, error: 'Password update failed anonymously.' });
                }
            });

        } catch (hashErr) {
            console.error("[UPDATE ERROR] Encrypt failed:", hashErr);
            res.redirect('/forgot-password?error=Encryption+Error');
        }
    });
});
// =============================================================

// 🔒 SECURITY FIX: Setup routes only in development
// Temporary route to reset admin password to 'admin123'
app.get('/setup-admin', async (req, res) => {
    // ✅ Only allow in development
    if (process.env.NODE_ENV === 'production') {
        return res.status(404).send('Not Found');
    }

    try {
        const hashedPassword = await bcrypt.hash('admin123', 10);
        // Check if admin exists
        db.execute('SELECT * FROM users WHERE username = "admin"', (err, results) => {
            if (results && results.length > 0) {
                // Update existing admin
                db.execute('UPDATE users SET password = ? WHERE username = "admin"', [hashedPassword], (err2) => {
                    if (err2) return res.send("Error updating admin: " + err2.message);
                    res.send("Admin password updated to 'admin123'. Try logging in now.");
                });
            } else {
                // Create new admin
                db.execute('INSERT INTO users (username, password, role, full_name, email) VALUES ("admin", ?, "admin", "System Administrator", "admin@waqar.edu.pk")', [hashedPassword], (err2) => {
                    if (err2) return res.send("Error creating admin: " + err2.message);
                    res.send("Admin created with password 'admin123'. Try logging in now.");
                });
            }
        });
    } catch (e) {
        res.send("Error: " + e.message);
    }
});

app.get('/setup-teacher', async (req, res) => {
    // ✅ Only allow in development
    if (process.env.NODE_ENV === 'production') {
        return res.status(404).send('Not Found');
    }

    try {
        const hashedPassword = await bcrypt.hash('teacher123', 10);
        db.query('SELECT id FROM users WHERE username = "teacher"', (err, results) => {
            if (results && results.length > 0) {
                db.execute('UPDATE users SET password = ? WHERE username = "teacher"', [hashedPassword], (err2) => {
                    res.send("Teacher 'teacher' password reset to 'teacher123'.");
                });
            } else {
                db.execute('INSERT INTO users (username, password, role, full_name, email) VALUES ("teacher", ?, "teacher", "Sample Teacher", "teacher@waqar.edu.pk")', [hashedPassword], (err2, result) => {
                    const userId = result.insertId;
                    db.execute('INSERT INTO teachers (user_id, subject, phone, salary) VALUES (?, "Mathematics", "1234567890", 50000)', [userId], (err3) => {
                        res.send("Teacher 'teacher' created with password 'teacher123'.");
                    });
                });
            }
        });
    } catch (e) {
        res.send("Error: " + e.message);
    }
});

app.get('/setup-student', async (req, res) => {
    // ✅ Only allow in development
    if (process.env.NODE_ENV === 'production') {
        return res.status(404).send('Not Found');
    }

    try {
        const hashedPassword = await bcrypt.hash('student123', 10);
        db.query('SELECT id FROM users WHERE username = "student"', (err, results) => {
            if (results && results.length > 0) {
                db.execute('UPDATE users SET password = ? WHERE username = "student"', [hashedPassword], (err2) => {
                    res.send("Student 'student' password reset to 'student123'.");
                });
            } else {
                db.execute('INSERT INTO users (username, password, role, full_name, email) VALUES ("student", ?, "student", "Ayaan Ali", "ayaan@waqar.edu.pk")', [hashedPassword], (err2, result) => {
                    const userId = result.insertId;
                    db.query('SELECT id FROM classes LIMIT 1', (err3, classes) => {
                        const classId = classes[0] ? classes[0].id : 1;
                        db.execute('INSERT INTO students (user_id, class_id, roll_no, phone) VALUES (?, ?, "101", "0987654321")', [userId, classId], (err4) => {
                            res.send("Student 'student' created with password 'student123'.");
                        });
                    });
                });
            }
        });
    } catch (e) {
        res.send("Error: " + e.message);
    }
});


app.get('/logout', (req, res) => {
    req.session.destroy();
    res.redirect('/login');
});

// Admin Dashboard
app.get('/admin/dashboard', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') {
        return res.redirect('/login');
    }

    const isSuperAdmin = (req.session.user.username === 'admin' && Number(req.session.user.campus_id) === 1);

    // 🛡️ PEOPLE SECURITY: Forced MFA for Super-Admin (Temporarily Disabled for Recovery)
    /*
    if (isSuperAdmin && !req.session.user.mfa_enabled) {
        return res.redirect('/admin/mfa-setup?force=true');
    }
    */

    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let statsQuery = '';
    let params = [];

    if (isSuperAdmin) {
        // SUPER ADMIN: Global Stats
        statsQuery = `
            SELECT 
                (SELECT COUNT(*) FROM students) as total_students,
                (SELECT COUNT(*) FROM teachers) as total_teachers,
                (SELECT SUM(remaining_balance) FROM vouchers WHERE remaining_balance > 0) as total_fees_pending
        `;
    } else {
        // REGULAR ADMIN: Campus Stats
        statsQuery = `
            SELECT 
                (SELECT COUNT(*) FROM students WHERE campus_id = ?) as total_students,
                (SELECT COUNT(*) FROM teachers WHERE campus_id = ?) as total_teachers,
                (SELECT SUM(remaining_balance) FROM vouchers WHERE remaining_balance > 0 AND campus_id = ?) as total_fees_pending
        `;
        params = [campusId, campusId, campusId];
    }

    db.query(statsQuery, params, (err, results) => {
        if (err) {
            console.error(err);
            return res.status(500).send('Database Error: ' + err.message);
        }
        const data = results[0];
        const newUser = req.session.new_user || null;
        if (newUser) delete req.session.new_user;

        // 📱 MOBILE QUICK LOGIN QR
        const protocol = req.protocol;
        const host = req.get('host');
        const campusCode = req.session.campus ? req.session.campus.campus_code : 'MAIN';
        const loginUrl = `${protocol}://${host}/login?campus_code=${campusCode}`;

        QRCode.toDataURL(loginUrl, (err, qrCodeUrl) => {
            res.render('admin/dashboard', {
                total_students: data.total_students || 0,
                total_teachers: data.total_teachers || 0,
                total_fees_pending: data.total_fees_pending || 0,
                newUser: newUser,
                isSuperAdmin: isSuperAdmin,
                qrCode: qrCodeUrl || null,
                loginUrl: loginUrl
            });
        });
    });
});

// Admin - Manage Students (Class-wise Blocks + Global Search)
app.get('/admin/students', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const classId = req.query.class_id || null;
    let searchQuery = req.query.search || null;
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    if (searchQuery) {
        // 🔒 SECURITY: Input Validation & Sanitization
        try {
            searchQuery = searchQuery.trim();

            // Validate minimum length
            if (searchQuery.length < 2) {
                console.log(`[STUDENT SEARCH] Query too short: "${searchQuery}"`);
                return res.redirect('/admin/students');
            }

            // Validate maximum length (prevent DoS)
            if (searchQuery.length > 100) {
                console.log(`[STUDENT SEARCH] Query too long: ${searchQuery.length} chars`);
                return res.redirect('/admin/students');
            }

            // Search Mode
            console.log(`[STUDENT SEARCH] Searching for: "${searchQuery}" (SuperAdmin: ${isSuperAdmin})`);

            let searchSql = '';
            let searchParams = [];
            const searchPattern = `%${searchQuery}%`;

            if (isSuperAdmin) {
                // Global Search
                searchSql = `
                    SELECT s.*, u.full_name, u.username, u.email, c.class_name, c.section, cp.campus_name 
                    FROM students s 
                    JOIN users u ON s.user_id = u.id 
                    JOIN classes c ON s.class_id = c.id
                    LEFT JOIN campuses cp ON s.campus_id = cp.id
                    WHERE (u.full_name LIKE ? OR u.username LIKE ? OR s.roll_no LIKE ?)
                `;
                searchParams = [searchPattern, searchPattern, searchPattern];
            } else {
                // Campus Search
                searchSql = `
                    SELECT s.*, u.full_name, u.username, u.email, c.class_name, c.section 
                    FROM students s 
                    JOIN users u ON s.user_id = u.id 
                    JOIN classes c ON s.class_id = c.id
                    WHERE s.campus_id = ? AND (u.full_name LIKE ? OR u.username LIKE ? OR s.roll_no LIKE ?)
                `;
                searchParams = [campusId, searchPattern, searchPattern, searchPattern];
            }

            db.query(searchSql, searchParams, (err, students) => {
                if (err) {
                    console.error('[STUDENT SEARCH ERROR]', err);
                    return res.redirect('/admin/students?error=Search failed. Please try again.');
                }

                // Fetch classes for dropdowns (Scoped or Global?) -> Usually for "Add Student" modal
                // We'll just fetch relevant classes
                const classQuery = isSuperAdmin ? 'SELECT * FROM classes' : 'SELECT * FROM classes WHERE campus_id = ?';
                const classParams = isSuperAdmin ? [] : [campusId];

                db.query(classQuery, classParams, (err2, classes) => {
                    if (err2) {
                        console.error('[CLASS FETCH ERROR]', err2);
                        return res.redirect('/admin/students?error=Failed to load classes.');
                    }

                    res.render('admin/students', {
                        students: students || [],
                        classes: classes || [],
                        currentClass: { class_name: 'Search Results', section: `"${searchQuery}"` },
                        viewMode: 'students',
                        selectedClass: null,
                        isSuperAdmin
                    });
                });
            });
        } catch (error) {
            console.error('[STUDENT SEARCH EXCEPTION]', error);
            return res.redirect('/admin/students?error=Search error occurred.');
        }

    } else if (!classId) {
        // View Mode: Blocks
        let query = '';
        let params = [];

        if (isSuperAdmin) {
            query = `
                SELECT c.*, COUNT(s.id) as student_count, cp.campus_name 
                FROM classes c 
                LEFT JOIN students s ON c.id = s.class_id 
                LEFT JOIN campuses cp ON c.campus_id = cp.id
                GROUP BY c.id
            `;
        } else {
            query = `
                SELECT c.*, COUNT(s.id) as student_count 
                FROM classes c 
                LEFT JOIN students s ON c.id = s.class_id AND s.campus_id = ?
                WHERE c.campus_id = ?
                GROUP BY c.id
            `;
            params = [campusId, campusId];
        }

        db.query(query, params, (err, classes) => {
            if (err) console.error(err);
            res.render('admin/students', { classes, students: [], viewMode: 'classes', isSuperAdmin });
        });
    } else {
        // View Mode: Student List in Class
        // For Super Admin, we still show the class even if it's in another campus (accessible via ID)
        const studentQuery = `
            SELECT s.*, u.full_name, u.username, u.email, c.class_name, c.section 
            FROM students s 
            JOIN users u ON s.user_id = u.id 
            JOIN classes c ON s.class_id = c.id
            WHERE s.class_id = ?
        `; // No campus filter needed here as class_id implies it, but we should secure it for regular admin

        let secureQuery = studentQuery;
        let secureParams = [classId];

        if (!isSuperAdmin) {
            secureQuery += ' AND s.campus_id = ?';
            secureParams.push(campusId);
        }

        db.query(secureQuery, secureParams, (err, students) => {
            const classQuery = isSuperAdmin ? 'SELECT * FROM classes' : 'SELECT * FROM classes WHERE campus_id = ?';
            const classParams = isSuperAdmin ? [] : [campusId];

            db.query(classQuery, classParams, (err2, classes) => {
                const currentClass = classes.find(c => c.id == classId) || { class_name: 'Unknown', section: '' };
                res.render('admin/students', {
                    students,
                    classes,
                    currentClass,
                    viewMode: 'students',
                    selectedClass: classId,
                    isSuperAdmin
                });
            });
        });
    }
});

// Admin - Add Student (Strict Campus Scope)
// 🔒 CSRF Protected
app.post('/admin/students/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    // 1. Inputs with 🔒 XSS Sanitization
    const full_name = xss(req.body.full_name || '').trim();
    const username = xss(req.body.username || '').trim();
    const class_id = req.body.class_id;
    const roll_no = xss(req.body.roll_no || '').trim();
    const father_name = xss(req.body.father_name || '').trim();
    const phone = xss(req.body.phone || '').trim();

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    if (!full_name || !username || !class_id) {
        return res.redirect(`/admin/students?error=${encodeURIComponent('Missing required fields: Name, Username, and Class are required.')}`);
    }

    // 2. Validate Class Ownership (Security)
    // If regular admin, ensure class belongs to their campus
    const clsCheck = isSuperAdmin ? 'SELECT id, campus_id FROM classes WHERE id = ?' : 'SELECT id, campus_id FROM classes WHERE id = ? AND campus_id = ?';
    const clsParams = isSuperAdmin ? [class_id] : [class_id, campusId];

    db.query(clsCheck, clsParams, async (err, clsResult) => {
        if (err || clsResult.length === 0) {
            return res.redirect(`/admin/students?error=${encodeURIComponent('Invalid Class selected or Access Denied.')}`);
        }

        const targetCampusId = clsResult[0].campus_id; // Store student in the class's campus

        // 3. Create User Account
        try {
            // Default password logic (e.g., username123)
            const passwordHash = await bcrypt.hash(username + '123', 10);

            // Check username uniqueness (locally or globally?)
            // Schema has composite unique (username, campus_id). 
            db.execute('INSERT INTO users (username, password, role, full_name, campus_id) VALUES (?, ?, "student", ?, ?)',
                [username, passwordHash, full_name, targetCampusId], (err2, userRes) => {
                    if (err2) {
                        console.error("User Create Error:", err2);
                        // Redirect with specific error about username
                        return res.redirect(`/admin/students?error=${encodeURIComponent('Username "' + username + '" is already taken in this campus. Please use a unique username.')}&class_id=${class_id}`);
                    }

                    // 4. Create Student Record
                    db.execute('INSERT INTO students (user_id, class_id, roll_no, father_name, phone, campus_id) VALUES (?, ?, ?, ?, ?, ?)',
                        [userRes.insertId, class_id, roll_no, father_name, phone, targetCampusId], (err3) => {
                            if (err3) {
                                console.error("Student Create Error:", err3);
                                // Cleanup created user record to prevent orphans
                                db.execute('DELETE FROM users WHERE id = ?', [userRes.insertId]);
                                return res.redirect(`/admin/students?error=${encodeURIComponent('Error creating student profile: ' + err3.message)}&class_id=${class_id}`);
                            }
                            res.redirect(`/admin/students?class_id=${class_id}&success=${encodeURIComponent('Student added successfully!')}`);
                        });
                });
        } catch (e) {
            console.error(e);
            res.redirect(`/admin/students?error=${encodeURIComponent('System Error: ' + e.message)}`);
        }
    });
});

// 📥 DOWNLOAD STUDENT CSV TEMPLATE
app.get('/admin/students/import/template', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const csvContent = "full_name,username,password,roll_no,class_name,section,father_name,phone,email\n" +
        "Ali Ahmed,ali101,pass123,101,three,A,Ahmed Khan,0300-1112223,ali@example.com\n" +
        "Sara Khan,sara102,pass123,102,5th,B,Khan Sahab,0300-4445556,sara@example.com\n";

    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename=student_import_template.csv');
    res.status(200).send(csvContent);
});

// 📥 BULK STUDENT CSV IMPORT
// 🔒 SECURITY FIX: Robust CSV upload validation
const csvUpload = multer({
    dest: 'uploads/temp/',
    limits: {
        fileSize: 5 * 1024 * 1024 // Increased to 5MB for larger cohorts
    },
    fileFilter: (req, file, cb) => {
        const allowedExtensions = ['.csv'];
        const allowedMimes = [
            'text/csv',
            'application/vnd.ms-excel',
            'text/plain',
            'application/csv',
            'text/comma-separated-values',
            'application/octet-stream' // 🚩 Added to handle files without extensions or generic binary streams
        ];

        const extension = path.extname(file.originalname).toLowerCase();
        const mimeType = file.mimetype;

        console.log(`[CSV ATTEMPT] File: ${file.originalname} | Detected MIME: ${mimeType} | Ext: ${extension}`);

        if (allowedExtensions.includes(extension) || allowedMimes.includes(mimeType)) {
            cb(null, true);
        } else {
            console.warn(`[SECURITY] Rejected file upload: ${file.originalname} (MIME: ${mimeType})`);
            cb(new Error('Only CSV files are allowed! Please ensure your file has a .csv extension and is saved in Comma Separated format.'));
        }
    }
});
// 🔒 CSRF Protected
app.post('/admin/students/import/csv', csvUpload.single('csvFile'), csrfProtection, async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    if (!req.file) {
        console.log('[CSV IMPORT] No file uploaded');
        return res.redirect('/admin/students?error=No file uploaded');
    }

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);
    const results = [];
    const filePath = req.file.path;

    console.log(`[CSV IMPORT] Starting import for ${isSuperAdmin ? 'Super Admin' : 'Campus Admin'} (Campus ID: ${campusId})`);
    console.log(`[CSV IMPORT] File path: ${filePath}`);

    fs.createReadStream(filePath)
        .pipe(csvParser())
        .on('data', (data) => {
            console.log('[CSV IMPORT] Row parsed:', data);
            results.push(data);
        })
        .on('end', async () => {
            console.log(`[CSV IMPORT] CSV parsing complete. Total rows: ${results.length}`);

            // 🔒 SECURITY FIX: Check if CSV is empty
            if (results.length === 0) {
                if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
                console.log('[CSV IMPORT] CSV file is empty');
                return res.redirect('/admin/students?error=CSV file is empty or invalid format');
            }

            // 🔒 SECURITY FIX: Limit number of rows to prevent DOS
            if (results.length > 1000) {
                if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
                console.log(`[CSV IMPORT] CSV file too large: ${results.length} rows`);
                return res.redirect('/admin/students?error=CSV file too large (max 1000 rows)');
            }

            // Processing
            let successCount = 0;
            let errorCount = 0;
            const errors = [];

            db.getConnection(async (err, connection) => {
                if (err) {
                    console.error('[CSV IMPORT] Database connection failed:', err);
                    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
                    return res.redirect('/admin/students?error=Database connection failed');
                }

                console.log('[CSV IMPORT] Database connection established');

                try {
                    await connection.promise().beginTransaction();
                    console.log('[CSV IMPORT] Transaction started');

                    for (let i = 0; i < results.length; i++) {
                        const row = results[i];
                        console.log(`[CSV IMPORT] Processing row ${i + 1}/${results.length}:`, row);

                        try {
                            const { full_name, username, password, roll_no, class_name, section, father_name, phone, email } = row;

                            // 🔒 Validation (Human Friendly)
                            if (!full_name || !username || !password || !class_name || !section) {
                                const missing = [];
                                if (!full_name) missing.push('Full Name');
                                if (!username) missing.push('Username');
                                if (!password) missing.push('Password');
                                if (!class_name) missing.push('Class Name');
                                if (!section) missing.push('Section');
                                throw new Error(`Missing required fields: ${missing.join(', ')}`);
                            }

                            // 1. Resolve Class ID from Name + Section + Campus
                            console.log(`[CSV IMPORT] Looking up class: ${class_name} (${section}) for Campus: ${campusId}`);

                            // 🚀 ROBUST MATCHING: Ignore spaces and case (e.g., "grade 12" matches "grade12")
                            const [classes] = await connection.promise().query(
                                'SELECT id, campus_id FROM classes WHERE REPLACE(LOWER(class_name), " ", "") = REPLACE(LOWER(?), " ", "") AND LOWER(TRIM(section)) = LOWER(TRIM(?)) AND campus_id = ?',
                                [class_name, section, campusId]
                            );

                            if (classes.length === 0) {
                                throw new Error(`Class "${class_name} (${section})" not found in your campus database.`);
                            }

                            const resolvedClassId = classes[0].id;
                            const targetCampusId = classes[0].campus_id;

                            // 2. Create User Account
                            console.log(`[CSV IMPORT] Creating user: ${username}`);
                            const passwordHash = await bcrypt.hash(password, 10);

                            const [userRes] = await connection.promise().execute(
                                'INSERT INTO users (username, password, role, full_name, campus_id, email) VALUES (?, ?, "student", ?, ?, ?)',
                                [username, passwordHash, full_name, targetCampusId, email || null]
                            );

                            // 3. Create Student Profile
                            await connection.promise().execute(
                                'INSERT INTO students (user_id, class_id, roll_no, father_name, phone, campus_id) VALUES (?, ?, ?, ?, ?, ?)',
                                [userRes.insertId, resolvedClassId, roll_no || null, father_name || null, phone || null, targetCampusId]
                            );

                            console.log(`[CSV IMPORT] Success: ${username} imported to ${class_name}-${section}`);
                            successCount++;
                        } catch (e) {
                            errorCount++;
                            const errorMsg = `${row.username || 'Row ' + (i + 1)}: ${e.message}`;
                            errors.push(errorMsg);
                            console.error(`[CSV IMPORT] Row ${i + 1} Error:`, e.message);
                        }
                    }

                    console.log(`[CSV IMPORT] Processing complete. Success: ${successCount}, Errors: ${errorCount}`);

                    await connection.promise().commit();
                    console.log('[CSV IMPORT] Transaction committed');
                    connection.release();
                    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);

                    const msg = `Imported ${successCount} students successfully. ${errorCount > 0 ? `Failed: ${errorCount}` : ''}`;
                    const errorDetails = errors.length > 0 ? '&errors=' + encodeURIComponent(errors.slice(0, 5).join(' | ')) : '';
                    console.log(`[CSV IMPORT] Redirecting with message: ${msg}`);
                    res.redirect(`/admin/students?success=${encodeURIComponent(msg)}${errorDetails}`);

                } catch (transactionErr) {
                    console.error('[CSV IMPORT] Transaction error:', transactionErr);
                    await connection.promise().rollback();
                    console.log('[CSV IMPORT] Transaction rolled back');
                    connection.release();
                    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
                    res.redirect('/admin/students?error=' + encodeURIComponent('Bulk import failed: ' + transactionErr.message));
                }
            });
        })
        .on('error', (parseError) => {
            console.error('[CSV IMPORT] CSV parsing error:', parseError);
            if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
            res.redirect('/admin/students?error=' + encodeURIComponent('CSV parsing failed: ' + parseError.message));
        });
});

// Admin - Delete Student (Strict Campus Verification)
app.get('/admin/students/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // 1. Find Student & User ID
    let q = 'SELECT id, user_id FROM students WHERE id = ?';
    let p = [req.params.id];

    if (!isSuperAdmin) {
        q += ' AND campus_id = ?';
        p.push(campusId);
    }

    db.query(q, p, (err, results) => {
        if (err || results.length === 0) return res.redirect('/admin/students');

        const student = results[0];

        // 2. Delete User (Cascade should handle student, but let's be safe)
        // If we delete User, foreign key in students (user_id) should cascade delete the student row?
        // Let's delete student first then user to be sure, or rely on schema.
        // Safer: Delete User.
        db.execute('DELETE FROM users WHERE id = ?', [student.user_id], (err2) => {
            if (err2) console.error(err2);
            // Also explicitly delete student if cascade fails
            db.execute('DELETE FROM students WHERE id = ?', [student.id], () => {
                res.redirect('/admin/students');
            });
        });
    });
});


// Admin - Manage Teachers
app.get('/admin/teachers', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = '';
    let params = [];

    if (isSuperAdmin) {
        query = `
            SELECT t.*, t.salary, u.full_name, u.username, u.email, cp.campus_name 
            FROM teachers t 
            JOIN users u ON t.user_id = u.id
            LEFT JOIN campuses cp ON t.campus_id = cp.id
        `;
    } else {
        query = `
            SELECT t.*, t.salary, u.full_name, u.username, u.email 
            FROM teachers t 
            JOIN users u ON t.user_id = u.id
            WHERE t.campus_id = ?
        `;
        params = [campusId];
    }

    db.query(query, params, (err, teachers) => {
        if (err) return res.send("Error fetching teachers");
        res.render('admin/teachers', { teachers, isSuperAdmin });
    });
});

// Admin - Add Teacher (Strict Campus Scope)
app.post('/admin/teachers/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const { full_name, username, password, subject, phone, salary, email } = req.body;
    const isSuperAdmin = (req.session.user.username === 'admin' && Number(req.session.user.campus_id) === 1);

    // For creation, we default to the currently active campus context
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    if (!full_name || !username || !password) {
        return res.status(400).send("Missing required fields: name, username, and password are required");
    }

    // 💰 VALIDATION: Salary Range Check (Relaxed to allow lower/blank salaries)
    const salaryNum = parseFloat(salary) || 0;
    if (salary && salaryNum > 1000000) {
        return res.status(400).send("Error: Maximum salary is Rs. 1,000,000. Please verify the amount.");
    }

    // 1. Create User
    (async () => {
        try {
            // Use the password provided in the form
            const passwordHash = await bcrypt.hash(password, 10);

            db.execute('INSERT INTO users (username, password, role, full_name, email, campus_id) VALUES (?, ?, "teacher", ?, ?, ?)',
                [username, passwordHash, full_name, email, campusId], (err, userRes) => {
                    if (err) {
                        console.error(err);
                        return res.send("Error creating user/username token");
                    }

                    // 2. Create Teacher Profile
                    db.execute('INSERT INTO teachers (user_id, subject, phone, salary, campus_id) VALUES (?, ?, ?, ?, ?)',
                        [userRes.insertId, subject, phone, salaryNum, campusId], (err2) => {
                            if (err2) {
                                console.error("Teacher Create Error:", err2);
                                // Cleanup created user record to prevent orphans
                                db.execute('DELETE FROM users WHERE id = ?', [userRes.insertId]);
                                return res.status(500).send("Error creating teacher profile: " + err2.message);
                            }
                            res.redirect('/admin/teachers');
                        });
                });
        } catch (e) {
            console.error(e);
            res.send("Server Error");
        }
    })();
});

// Admin - Delete Teacher (Strict Campus Verification)
app.get('/admin/teachers/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // 1. Find Teacher & User ID
    let q = 'SELECT id, user_id FROM teachers WHERE id = ?';
    let p = [req.params.id];

    if (!isSuperAdmin) {
        q += ' AND campus_id = ?';
        p.push(campusId);
    }

    db.query(q, p, (err, results) => {
        if (err || results.length === 0) return res.redirect('/admin/teachers');

        const teacher = results[0];

        // 2. Delete User (Cascade to Teacher)
        db.execute('DELETE FROM users WHERE id = ?', [teacher.user_id], (err2) => {
            if (err2) console.error(err2);
            // Explicit delete just in case
            db.execute('DELETE FROM teachers WHERE id = ?', [teacher.id], () => {
                res.redirect('/admin/teachers');
            });
        });
    });
});

// Admin - Reset User Password
app.post('/admin/users/reset-password', csrfProtection, async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const { user_id, new_password, return_url } = req.body;
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    if (!user_id || !new_password) {
        return res.redirect(return_url || '/admin/dashboard');
    }

    try {
        // Security: Verify user belongs to admin's campus (unless super admin)
        let verifyQuery = 'SELECT id, campus_id FROM users WHERE id = ?';
        let verifyParams = [user_id];

        if (!isSuperAdmin) {
            verifyQuery += ' AND campus_id = ?';
            verifyParams.push(campusId);
        }

        db.query(verifyQuery, verifyParams, async (err, users) => {
            if (err || users.length === 0) {
                console.error("Password Reset Error: User not found or access denied");
                return res.redirect(return_url || '/admin/dashboard');
            }

            // Hash new password
            const hashedPassword = await bcrypt.hash(new_password, 10);

            // Update password
            db.execute('UPDATE users SET password = ? WHERE id = ?', [hashedPassword, user_id], (updateErr) => {
                if (updateErr) {
                    console.error("Password Update Error:", updateErr);
                }
                res.redirect(return_url || '/admin/dashboard');
            });
        });
    } catch (e) {
        console.error("Password Reset Exception:", e);
        res.redirect(return_url || '/admin/dashboard');
    }
});


// Admin - Manage Classes
app.get('/admin/classes', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = isSuperAdmin ?
        'SELECT c.*, cp.campus_name, u.full_name as teacher_name FROM classes c LEFT JOIN campuses cp ON c.campus_id = cp.id LEFT JOIN teachers t ON c.class_teacher_id = t.id LEFT JOIN users u ON t.user_id = u.id ORDER BY c.class_name, c.section' :
        'SELECT c.*, u.full_name as teacher_name FROM classes c LEFT JOIN teachers t ON c.class_teacher_id = t.id LEFT JOIN users u ON t.user_id = u.id WHERE c.campus_id = ? ORDER BY c.class_name, c.section';
    const params = isSuperAdmin ? [] : [campusId];

    db.query(query, params, (err, classes) => {
        if (err) {
            console.error(err);
            return res.render('admin/classes', { classes: [], teachers: [], isSuperAdmin });
        }

        // Fetch teachers to assign as Class Teachers (Scoped)
        const teacherQuery = isSuperAdmin ?
            'SELECT t.id, u.full_name FROM teachers t JOIN users u ON t.user_id = u.id' :
            'SELECT t.id, u.full_name FROM teachers t JOIN users u ON t.user_id = u.id WHERE t.campus_id = ?';
        const teacherParams = isSuperAdmin ? [] : [campusId];

        db.query(teacherQuery, teacherParams, (err2, teachers) => {
            res.render('admin/classes', {
                classes: classes || [],
                teachers: teachers || [],
                isSuperAdmin
            });
        });
    });
});

app.post('/admin/classes/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { class_name, section, monthly_fee } = req.body;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    db.query('INSERT INTO classes (class_name, section, monthly_fee, campus_id) VALUES (?, ?, ?, ?)',
        [class_name, section, monthly_fee, campusId], (err) => {
            if (err) console.error(err);
            res.redirect('/admin/classes');
        });
});

app.post('/admin/classes/update-fee', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { class_id, monthly_fee } = req.body;
    db.query('UPDATE classes SET monthly_fee = ? WHERE id = ?', [monthly_fee, class_id], (err) => {
        if (err) console.error(err);
        res.redirect('/admin/classes');
    });
});

// Assign Class Teacher
app.post('/admin/classes/assign-teacher', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { class_id, teacher_id } = req.body;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // Verify ownership if not super admin
    const checkQuery = req.session.user.campus_id === 1 ? 'SELECT id FROM classes WHERE id = ?' : 'SELECT id FROM classes WHERE id = ? AND campus_id = ?';
    const checkParams = req.session.user.campus_id === 1 ? [class_id] : [class_id, campusId];

    db.query(checkQuery, checkParams, (err, results) => {
        if (err || results.length === 0) return res.send("Access Denied or Invalid Class");

        db.execute('UPDATE classes SET class_teacher_id = ? WHERE id = ?', [teacher_id || null, class_id], (err2) => {
            if (err2) console.error(err2);
            res.redirect('/admin/classes?success=Teacher+Assigned');
        });
    });
});

app.get('/admin/classes/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = 'DELETE FROM classes WHERE id = ?';
    let params = [req.params.id];

    if (!isSuperAdmin) {
        query += ' AND campus_id = ?';
        params.push(campusId);
    }

    db.query(query, params, (err) => {
        if (err) console.error(err);
        res.redirect('/admin/classes');
    });
});


// Admin - Fees Dashboard (Scoped to Campus)
app.get('/admin/fees', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = '';
    let params = [];

    if (isSuperAdmin) {
        query = `
            SELECT v.*, u.full_name, c.class_name, c.section, cp.campus_name
            FROM vouchers v
            JOIN students s ON v.student_id = s.id
            JOIN users u ON s.user_id = u.id
            JOIN classes c ON s.class_id = c.id
            LEFT JOIN campuses cp ON v.campus_id = cp.id
            ORDER BY v.id DESC
        `;
    } else {
        query = `
            SELECT v.*, u.full_name, c.class_name, c.section
            FROM vouchers v
            JOIN students s ON v.student_id = s.id
            JOIN users u ON s.user_id = u.id
            JOIN classes c ON s.class_id = c.id
            WHERE v.campus_id = ?
            ORDER BY v.id DESC
        `;
        params = [campusId];
    }

    db.query(query, params, (err, vouchers) => {
        if (err) {
            console.error("Fees Query Error:", err);
            vouchers = [];
        }

        // Students dropdown for "Add Individual Fee" - needs to follow same logic
        let studentSql = '';
        let studentParams = [];

        if (isSuperAdmin) {
            studentSql = 'SELECT s.id, u.full_name, cp.campus_name FROM students s JOIN users u ON s.user_id = u.id LEFT JOIN campuses cp ON s.campus_id = cp.id';
        } else {
            studentSql = 'SELECT s.id, u.full_name FROM students s JOIN users u ON s.user_id = u.id WHERE s.campus_id = ?';
            studentParams = [campusId];
        }

        db.query(studentSql, studentParams, (err2, students) => {
            if (err2) {
                console.error("Fees Student Query Error:", err2);
                students = [];
            }
            res.render('admin/fees', { vouchers: vouchers || [], students: students || [], isSuperAdmin });
        });
    });
});

// 📊 EXPORT FEE LEDGER TO EXCEL
// 🔒 SECURITY FIX: Rate limited to prevent abuse
app.get('/admin/fees/export/excel', exportLimiter, async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    try {
        let query = '';
        let params = [];

        if (isSuperAdmin) {
            query = `
                SELECT v.*, u.full_name, c.class_name, c.section, cp.campus_name
                FROM vouchers v
                JOIN students s ON v.student_id = s.id
                JOIN users u ON s.user_id = u.id
                JOIN classes c ON s.class_id = c.id
                LEFT JOIN campuses cp ON v.campus_id = cp.id
                ORDER BY v.id DESC
            `;
        } else {
            query = `
                SELECT v.*, u.full_name, c.class_name, c.section
                FROM vouchers v
                JOIN students s ON v.student_id = s.id
                JOIN users u ON s.user_id = u.id
                JOIN classes c ON s.class_id = c.id
                WHERE v.campus_id = ?
                ORDER BY v.id DESC
            `;
            params = [campusId];
        }

        db.query(query, params, async (err, vouchers) => {
            if (err) {
                console.error('[EXPORT ERROR]', err);
                return res.status(500).send('Error fetching fee data');
            }

            // Create Excel Workbook
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Fee Ledger');

            // Set column headers
            const columns = [
                { header: 'Voucher ID', key: 'id', width: 12 },
                { header: 'Student Name', key: 'full_name', width: 25 },
                { header: 'Class', key: 'class_name', width: 15 },
                { header: 'Section', key: 'section', width: 12 },
                { header: 'Academic Year', key: 'academic_year', width: 15 },
                { header: 'Total Charged (Rs)', key: 'total_charged', width: 18 },
                { header: 'Total Paid (Rs)', key: 'total_paid', width: 18 },
                { header: 'Balance Due (Rs)', key: 'remaining_balance', width: 18 }
            ];

            if (isSuperAdmin) {
                columns.splice(4, 0, { header: 'Campus', key: 'campus_name', width: 20 });
            }

            worksheet.columns = columns;

            // Style header row
            worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            worksheet.getRow(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF3B82F6' }
            };
            worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

            // Add data rows
            vouchers.forEach(v => {
                const row = {
                    id: v.id,
                    full_name: v.full_name,
                    class_name: v.class_name,
                    section: v.section,
                    academic_year: v.academic_year,
                    total_charged: parseFloat(v.total_charged),
                    total_paid: parseFloat(v.total_paid),
                    remaining_balance: parseFloat(v.remaining_balance)
                };

                if (isSuperAdmin) {
                    row.campus_name = v.campus_name || '-';
                }

                worksheet.addRow(row);
            });

            // Add totals row
            const totalRow = worksheet.addRow({
                id: '',
                full_name: 'TOTAL',
                class_name: '',
                section: '',
                academic_year: '',
                total_charged: vouchers.reduce((sum, v) => sum + parseFloat(v.total_charged), 0),
                total_paid: vouchers.reduce((sum, v) => sum + parseFloat(v.total_paid), 0),
                remaining_balance: vouchers.reduce((sum, v) => sum + parseFloat(v.remaining_balance), 0)
            });

            totalRow.font = { bold: true };
            totalRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF3F4F6' }
            };

            // Format currency columns
            worksheet.getColumn('total_charged').numFmt = '#,##0.00';
            worksheet.getColumn('total_paid').numFmt = '#,##0.00';
            worksheet.getColumn('remaining_balance').numFmt = '#,##0.00';

            // Set response headers
            const filename = `Fee_Ledger_${new Date().toISOString().split('T')[0]}.xlsx`;
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

            // Write to response
            await workbook.xlsx.write(res);
            res.end();

            console.log(`[EXPORT SUCCESS] Fee ledger exported: ${filename}`);
        });
    } catch (error) {
        console.error('[EXPORT EXCEPTION]', error);
        res.status(500).send('Export failed');
    }
});



// Master Voucher Receipt (Annual Statement) - 🔒 IDOR Protection
app.get('/receipt/:id', (req, res) => {
    if (!req.session.user) return res.redirect('/login');

    const query = `
        SELECT v.*, u.full_name, c.class_name, c.section, s.roll_no, s.user_id, v.campus_id
        FROM vouchers v
        JOIN students s ON v.student_id = s.id
        JOIN users u ON s.user_id = u.id
        JOIN classes c ON s.class_id = c.id
        WHERE v.id = ?
    `;

    db.query(query, [req.params.id], (err, results) => {
        if (err || results.length === 0) return res.status(404).send("Receipt not found.");

        const voucher = results[0];
        const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);

        // 🔒 SECURITY FIX: Verify ownership/access
        if (req.session.user.role === 'student') {
            // Students can only view their own receipts
            if (voucher.user_id !== req.session.user.id) {
                return res.status(403).send("Unauthorized: You can only view your own receipts.");
            }
        } else if (req.session.user.role === 'admin' && !isSuperAdmin) {
            // Regular admins can only view receipts from their campus
            const campusId = req.session.campus ? req.session.campus.id : req.session.user.campus_id;
            if (voucher.campus_id !== campusId) {
                return res.status(403).send("Unauthorized: Access denied to this campus data.");
            }
        }
        // Super admin can view all receipts (no restriction)

        // Fetch items and payments for the statement
        db.query('SELECT * FROM fees WHERE voucher_id = ?', [voucher.id], (errI, items) => {
            db.query('SELECT * FROM fee_payments WHERE voucher_id = ?', [voucher.id], (errP, payments) => {
                res.render('receipt', { voucher, items, payments });
            });
        });
    });
});

// Admin - Individual Payment Transaction Receipt - 🔒 IDOR Protection
app.get('/payment-receipt/:id', (req, res) => {
    if (!req.session.user) return res.redirect('/login');

    const query = `
        SELECT p.*, f.month, f.year, f.remaining_amount, f.campus_id,
               u.full_name, c.class_name, s.roll_no, s.user_id
        FROM fee_payments p
        JOIN fees f ON p.fee_id = f.id
        JOIN students s ON f.student_id = s.id
        JOIN users u ON s.user_id = u.id
        JOIN classes c ON s.class_id = c.id
        WHERE p.id = ?
    `;

    db.query(query, [req.params.id], (err, results) => {
        if (err || results.length === 0) return res.status(404).send("Payment record not found.");

        const payment = results[0];
        const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);

        // 🔒 SECURITY FIX: Verify ownership/access
        if (req.session.user.role === 'student') {
            if (payment.user_id !== req.session.user.id) {
                return res.status(403).send("Unauthorized: You can only view your own payment receipts.");
            }
        } else if (req.session.user.role === 'admin' && !isSuperAdmin) {
            const campusId = req.session.campus ? req.session.campus.id : req.session.user.campus_id;
            if (payment.campus_id !== campusId) {
                return res.status(403).send("Unauthorized: Access denied to this campus data.");
            }
        }

        res.render('payment-receipt', { payment });
    });
});


// Admin - Manually Add Single Fee (Fine/Other)
// 🔒 CSRF Protected
app.post('/admin/fees/add-single', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { student_id, amount, title, month, year } = req.body;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    if (!student_id || !amount) return res.send("<script>alert('Please select a student and amount'); window.history.back();</script>");

    // 1. Get or Create Voucher
    db.query('SELECT * FROM vouchers WHERE student_id = ? AND academic_year = ?', [student_id, year], (err, vResult) => {
        if (err) return res.send("Error checking voucher: " + err.message);

        const addToVoucher = (voucherId) => {
            // 2. Add Fee Item
            // Note: We use the 'status' column or add a 'title' column? 
            // The current schema doesn't have a 'title' column in 'fees' table, it relies on month/year.
            // We'll append the Title to the month field for now like "Fine (Oct)" or just use Month if standard.
            // Actually, best to just use the month/year logic or standard fee. 
            // Let's assume this is for "Extra" fees, maybe we should use a specific month like "Misc-Jan".
            // For now, we will just use the provided Month/Year.

            checkBilling(voucherId, student_id, amount, month, year, campusId, (success) => {
                if (success) {
                    recalculateVoucher(voucherId, () => {
                        res.redirect('/admin/fees?success=added');
                    });
                } else {
                    res.send("<script>alert('Error adding fee: Duplicate entry for this month/year or DB error.'); window.history.back();</script>");
                }
            });
        };

        if (vResult.length === 0) {
            db.execute('INSERT INTO vouchers (student_id, academic_year, total_charged, total_paid, remaining_balance, campus_id) VALUES (?, ?, 0, 0, 0, ?)',
                [student_id, year, campusId], (err2, res2) => {
                    if (err2) return res.send("Error creating voucher: " + err2.message);
                    addToVoucher(res2.insertId);
                });
        } else {
            addToVoucher(vResult[0].id);
        }
    });
});

// Helper for Add Single
function checkBilling(voucherId, studentId, amount, month, year, campusId, cb) {
    // Check if fee exists?
    // If we want to allow MULTIPLE fees per month (like Tuition + Fine), we need to relax the unique constraint or compose the month string.
    // The current schema has: UNIQUE INDEX student_month_year
    // So we can't add two entries for "October" "2024".
    // We will try to append a random suffix or expect the user to choose a different "Type" if we had one.
    // For now, let's just Try Insert.

    const lineTotal = parseFloat(amount); // Simplified for single add

    db.execute(
        'INSERT INTO fees (voucher_id, student_id, amount, carried_due, line_total, month, year, status, campus_id) VALUES (?, ?, ?, 0, ?, ?, ?, "unpaid", ?)',
        [voucherId, studentId, amount, lineTotal, month, year, campusId],
        (err) => {
            if (err) {
                console.error("Add Fee Error:", err);
                // If duplicate, maybe we append to the existing amount?
                if (err.code === 'ER_DUP_ENTRY') {
                    // Update existing
                    db.execute('UPDATE fees SET amount = amount + ?, line_total = line_total + ? WHERE student_id = ? AND month = ? AND year = ?',
                        [amount, amount, studentId, month, year], (errUp) => {
                            cb(!errUp);
                        });
                    return;
                }
                cb(false);
            } else {
                cb(true);
            }
        }
    );
}

// Admin - LEDGER-BASED SYNC (ONE VOUCHER PER YEAR)
// 🔒 CSRF Protected
app.post('/admin/fees/generate-all', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { month, year } = req.body;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const studentQuery = `
        SELECT s.id as student_id, COALESCE(c.monthly_fee, 0) as class_fee 
        FROM students s 
        LEFT JOIN classes c ON s.class_id = c.id
        WHERE s.campus_id = ?
    `;

    db.query(studentQuery, [campusId], (err, students) => {
        if (err) {
            console.error("Sync Error - Fetch Students:", err);
            return res.redirect('/admin/fees?error=db_fetch_students');
        }
        if (!students || students.length === 0) return res.redirect('/admin/fees?info=no_students_found');

        console.log(`[SYNC START] Processing ${students.length} students for ${month} ${year} at Campus ${campusId}`);

        let processed = 0;
        let hasResponded = false;

        const checkFinished = () => {
            processed++;
            if (processed >= students.length && !hasResponded) {
                hasResponded = true;
                console.log("[SYNC COMPLETE]");
                res.redirect('/admin/fees?success=synced');
            }
        };

        students.forEach(student => {
            // 1. GET OR CREATE Master Voucher
            db.query('SELECT * FROM vouchers WHERE student_id = ? AND academic_year = ?', [student.student_id, year], (errV, voucherRes) => {
                if (errV) {
                    console.error(`Sync Error - Get Voucher (Student ${student.student_id}):`, errV);
                    return checkFinished();
                }

                const processVoucherLine = (v) => {
                    if (!v) return checkFinished();

                    // 2. CHECK IF record exists for (student_id, month, year)
                    db.query('SELECT * FROM fees WHERE student_id = ? AND month = ? AND year = ?', [student.student_id, month, year], (errF, feeRes) => {
                        if (errF) {
                            console.error("Sync Error - Check Fee:", errF);
                            return checkFinished();
                        }

                        if (feeRes.length === 0) {
                            // 3. INSERT (New Month)
                            const carriedDue = parseFloat(v.remaining_balance || 0);
                            const currentFee = parseFloat(student.class_fee || 0);
                            // Avoid adding 0 amount fees if you don't want clutter, but ledger usually needs it.
                            // We will add it even if 0 to show continuity.
                            const lineTotal = currentFee + carriedDue; // This logic is slightly flawed for ledger line items (should just be fee), but following existing pattern.

                            // actually 'line_total' in a localized fee record usually means "Amount + Arrears" at that point in time.
                            // We will keep it.

                            db.execute(
                                'INSERT INTO fees (voucher_id, student_id, amount, carried_due, line_total, month, year, status, campus_id) VALUES (?, ?, ?, ?, ?, ?, ?, "unpaid", ?)',
                                [v.id, student.student_id, currentFee, carriedDue, lineTotal, month, year, campusId],
                                (errI) => {
                                    if (errI) console.error("Sync Error - Insert Fee:", errI);
                                    recalculateVoucher(v.id, () => checkFinished());
                                }
                            );
                        } else {
                            // 4. UPDATE EXISTING (Link or Update Amount)
                            const existingFee = feeRes[0];
                            const needsLink = (existingFee.voucher_id === null);
                            const currentFeeAmt = parseFloat(existingFee.amount || 0);
                            const targetFeeAmt = parseFloat(student.class_fee || 0);

                            // If the fee in DB is 0 but class has a fee, update it.
                            // If DB has a fee but class is 0 (maybe manual override?), don't overwrite unless logic demands.
                            // We'll only update if current is 0 and target > 0 (Missing fee)
                            const needsFeeUpdate = (currentFeeAmt === 0 && targetFeeAmt > 0);

                            if (needsLink || needsFeeUpdate) {
                                const newAmount = needsFeeUpdate ? targetFeeAmt : currentFeeAmt;
                                db.execute(
                                    'UPDATE fees SET voucher_id = ?, amount = ? WHERE id = ?',
                                    [v.id, newAmount, existingFee.id],
                                    () => {
                                        recalculateVoucher(v.id, () => checkFinished());
                                    }
                                );
                            } else {
                                checkFinished();
                            }
                        }
                    });
                };


                if (voucherRes.length === 0) {
                    // Create Master Voucher (ONLY ONCE)
                    db.execute('INSERT INTO vouchers (student_id, academic_year, total_charged, total_paid, remaining_balance, campus_id) VALUES (?, ?, 0, 0, 0, ?)',
                        [student.student_id, year, campusId], (errInsV, results) => {
                            if (errInsV) {
                                console.error("Sync Error - Create Voucher:", errInsV);
                                return checkFinished();
                            }
                            // Fetch newly created voucher to be safe
                            const newVoucherStub = {
                                id: results.insertId,
                                remaining_balance: 0
                            };
                            processVoucherLine(newVoucherStub);
                        });
                } else {
                    processVoucherLine(voucherRes[0]);
                }
            });
        });
    });
});



// Admin - Ledger-Based Voucher Payment (TRANSACTION-SAFE) - SECURED
app.post('/admin/fees/pay/:id', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const paymentAmount = parseFloat(req.body.payment_amount || 0);
    const voucherId = req.params.id;
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    if (paymentAmount <= 0) return res.send("<script>alert('Invalid amount'); window.history.back();</script>");

    // Get connection from pool first
    db.getConnection((connErr, connection) => {
        if (connErr) {
            console.error('Connection Error:', connErr);
            return res.send("System Busy. Try again.");
        }

        connection.beginTransaction(err => {
            if (err) {
                connection.release();
                return res.send("System Busy. Try again.");
            }

            // Lock record FOR UPDATE
            let query = 'SELECT * FROM vouchers WHERE id = ?';
            let params = [voucherId];

            if (!isSuperAdmin) {
                query += ' AND campus_id = ?';
                params.push(campusId);
            }
            query += ' FOR UPDATE';

            connection.query(query, params, (err, results) => {
                if (err || results.length === 0) {
                    return connection.rollback(() => {
                        connection.release();
                        res.send('Voucher not found or access denied');
                    });
                }
                const voucher = results[0];

                if (paymentAmount > voucher.remaining_balance) {
                    return connection.rollback(() => {
                        connection.release();
                        res.send("<script>alert('Error: Payment exceeds current balance!'); window.history.back();</script>");
                    });
                }

                const receiptNo = `RCP-${Date.now()}`;
                // 1. Log Payment
                connection.execute('INSERT INTO fee_payments (voucher_id, amount_paid, payment_method, receipt_no, recorded_by, fee_id, campus_id) VALUES (?, ?, ?, ?, ?, NULL, ?)',
                    [voucherId, paymentAmount, req.body.method || 'Cash', receiptNo, req.session.user.id, voucher.campus_id], (err2) => {
                        if (err2) {
                            console.error("Payment Insert Error:", err2);
                            return connection.rollback(() => {
                                connection.release();
                                res.send("Payment Record Failure: " + err2.message);
                            });
                        }

                        // 2. Finalize Voucher State
                        const newPaid = parseFloat(voucher.total_paid) + paymentAmount;
                        const newBal = parseFloat(voucher.remaining_balance) - paymentAmount;

                        connection.execute('UPDATE vouchers SET total_paid = ?, remaining_balance = ? WHERE id = ?',
                            [newPaid, newBal, voucherId], (err3) => {
                                if (err3) {
                                    console.error("Voucher Update Error:", err3);
                                    return connection.rollback(() => {
                                        connection.release();
                                        res.send("Voucher Update Failure: " + err3.message);
                                    });
                                }

                                connection.commit(err4 => {
                                    if (err4) {
                                        console.error("Commit Error:", err4);
                                        return connection.rollback(() => {
                                            connection.release();
                                            res.send("Transaction Commit Failure");
                                        });
                                    }
                                    connection.release();
                                    res.redirect('/admin/fees?success=paid');
                                });
                            });
                    });
            });
        });
    });
});


// Admin - Revert Master Voucher (Clear all payments) - SECURED
app.post('/admin/fees/revert/:id', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = 'SELECT * FROM vouchers WHERE id = ?';
    let params = [req.params.id];
    if (!isSuperAdmin) {
        query += ' AND campus_id = ?';
        params.push(campusId);
    }

    db.query(query, params, (err, results) => {
        if (!results || results.length === 0) return res.send("Voucher not found or access denied");
        const voucher = results[0];

        // Delete all payments associated with this voucher
        db.execute('DELETE FROM fee_payments WHERE voucher_id = ?', [voucher.id], () => {
            // Delete income entries from main school ledger
            // Note: Since payments against vouchers are not tied to specific fee_id in the ledger (related_fee_id=0), we might need a better way to find them, 
            // but for now we follow the user prompt's ledger logic.
            recalculateVoucher(voucher.id, () => {
                res.redirect('/admin/fees?success=reverted');
            });
        });
    });
});


// Admin - Delete Master Voucher (and all line items) - SECURED
app.get('/admin/fees/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = 'SELECT * FROM vouchers WHERE id = ?';
    let params = [req.params.id];
    if (!isSuperAdmin) {
        query += ' AND campus_id = ?';
        params.push(campusId);
    }

    db.query(query, params, (err, results) => {
        if (err || !results.length) return res.redirect('/admin/fees');
        const v = results[0];

        // Cascade delete: vouchers -> fees (line items) and payments
        db.execute('DELETE FROM fees WHERE voucher_id = ?', [v.id], () => {
            db.execute('DELETE FROM fee_payments WHERE voucher_id = ?', [v.id], () => {
                db.execute('DELETE FROM vouchers WHERE id = ?', [v.id], () => {
                    res.redirect('/admin/fees?success=deleted');
                });
            });
        });
    });
});




// Helper: Recalculate Voucher Ledger
function recalculateVoucher(voucherId, cb) {
    if (!voucherId) {
        if (cb) cb();
        return;
    }
    // 1. Calculate Total Charged (Sum of fees)
    db.query('SELECT SUM(amount) as total_charged FROM fees WHERE voucher_id = ?', [voucherId], (err, res1) => {
        const charged = (res1 && res1[0] && res1[0].total_charged) ? parseFloat(res1[0].total_charged) : 0;

        // 2. Calculate Total Paid (Sum of payments)
        db.query('SELECT SUM(amount_paid) as total_paid FROM fee_payments WHERE voucher_id = ?', [voucherId], (err2, res2) => {
            const paid = (res2 && res2[0] && res2[0].total_paid) ? parseFloat(res2[0].total_paid) : 0;
            const balance = charged - paid;

            // 3. Update Voucher
            db.execute('UPDATE vouchers SET total_charged = ?, total_paid = ?, remaining_balance = ? WHERE id = ?',
                [charged, paid, balance, voucherId],
                (err3) => {
                    if (err3) console.error("Recalculate Error:", err3);
                    if (cb) cb();
                });
        });
    });
}


// Admin - Exam Management
app.get('/admin/exams', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = isSuperAdmin ?
        'SELECT e.*, c.campus_name FROM exams e LEFT JOIN campuses c ON e.campus_id = c.id ORDER BY start_date DESC' :
        'SELECT * FROM exams WHERE campus_id = ? ORDER BY start_date DESC';
    const params = isSuperAdmin ? [] : [campusId];

    db.query(query, params, (err, exams) => {
        if (err) exams = [];
        res.render('admin/exams', { exams, isSuperAdmin });
    });
});

// Admin - Add Exam
app.post('/admin/exams/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { exam_name, start_date, end_date } = req.body;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    db.execute('INSERT INTO exams (exam_name, start_date, end_date, campus_id) VALUES (?, ?, ?, ?)',
        [exam_name, start_date, end_date, campusId],
        (err) => {
            if (err) console.error(err);
            res.redirect('/admin/exams');
        });
});

// Admin - Delete Exam - SECURED
app.get('/admin/exams/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = 'DELETE FROM exams WHERE id = ?';
    let params = [req.params.id];

    if (!isSuperAdmin) {
        query += ' AND campus_id = ?';
        params.push(campusId);
    }

    db.execute(query, params, (err) => {
        if (err) console.error(err);
        res.redirect('/admin/exams');
    });
});

// Admin - Exam Results View (Class Wise)
app.get('/admin/exams/:id/results', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const examId = req.params.id;
    const selectedClass = req.query.class_id || null;
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);

    db.query('SELECT * FROM exams WHERE id = ?', [examId], (err, examResult) => {
        if (err || examResult.length === 0) return res.redirect('/admin/exams');
        const exam = examResult[0];

        if (!selectedClass) {
            // Overview: Show all classes
            // For overview, we want to see stats per class.
            const overviewQuery = `
                SELECT 
                    c.id, c.class_name, c.section, cp.campus_name,
                    COUNT(DISTINCT s.id) as total_students,
                    COUNT(DISTINCT r.id) as results_count
                FROM classes c
                LEFT JOIN campuses cp ON c.campus_id = cp.id
                LEFT JOIN students s ON c.id = s.class_id
                LEFT JOIN exam_results r ON s.id = r.student_id AND r.exam_id = ?
                WHERE c.campus_id = ? 
                GROUP BY c.id
            `;
            // NOTE: Even Super Admin should only see classes relevant to this exam's campus 
            // because exams are usually campus-specific entities.
            // If the user wants to see cross-campus exams, they'd likely need a Global Exam ID which is complex.
            // We assume Exam is bound to one campus.

            db.query(overviewQuery, [examId, exam.campus_id], (err2, overview) => {
                if (err2) console.error(err2);
                res.render('admin/exam_results', {
                    exam,
                    overview: overview || [],
                    results: [],
                    selectedClass: null,
                    viewMode: 'overview',
                    isSuperAdmin
                });
            });
            return;
        }

        // Detail View: Show results for selected class
        db.query('SELECT * FROM classes WHERE id = ?', [selectedClass], (errClass, classResult) => {
            const currentClass = classResult[0];
            const resultsQuery = `
                SELECT r.*, s.roll_no, u.full_name
                FROM students s
                JOIN users u ON s.user_id = u.id
                LEFT JOIN exam_results r ON s.id = r.student_id AND r.exam_id = ?
                WHERE s.class_id = ?
                ORDER BY s.roll_no
            `;
            db.query(resultsQuery, [examId, selectedClass], (err3, results) => {
                if (err3) console.error(err3);
                res.render('admin/exam_results', {
                    exam,
                    overview: [],
                    results,
                    selectedClass,
                    currentClass, // Pass class details
                    viewMode: 'detail',
                    isSuperAdmin
                });
            });
        });
    });
});

// Admin - Add Exam Result
app.post('/admin/exams/results/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { exam_id, student_id, subject, marks_obtained, total_marks } = req.body;

    // Use student's campus
    db.query('SELECT campus_id FROM students WHERE id = ?', [student_id], (errS, resS) => {
        const campusId = (resS && resS.length) ? resS[0].campus_id : 1;

        // Calculate Grade
        const percentage = (marks_obtained / total_marks) * 100;
        let grade = 'F';
        if (percentage >= 90) grade = 'A+';
        else if (percentage >= 80) grade = 'A';
        else if (percentage >= 70) grade = 'B';
        else if (percentage >= 60) grade = 'C';
        else if (percentage >= 50) grade = 'D';

        const query = 'INSERT INTO exam_results (exam_id, student_id, subject, marks_obtained, total_marks, grade, campus_id) VALUES (?, ?, ?, ?, ?, ?, ?)';
        db.execute(query, [exam_id, student_id, subject, marks_obtained, total_marks, grade, campusId], (err) => {
            if (err) console.error(err);
            res.redirect(`/admin/exams/${exam_id}/results`);
        });
    });
});

// Admin - Attendance Management
app.get('/admin/attendance', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const selectedClass = req.query.class_id || null;
    const selectedDate = req.query.date || new Date().toISOString().split('T')[0];
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // Get all classes with attendance stats for today
    if (!selectedClass) {
        let overviewQuery = '';
        let params = [];

        if (isSuperAdmin) {
            overviewQuery = `
                SELECT 
                    c.id, c.class_name, c.section, cp.campus_name,
                    COUNT(DISTINCT s.id) as total_students,
                    COUNT(DISTINCT a.id) as marked_count,
                    SUM(CASE WHEN a.status = 'present' THEN 1 ELSE 0 END) as present_count,
                    SUM(CASE WHEN a.status = 'absent' THEN 1 ELSE 0 END) as absent_count,
                    SUM(CASE WHEN a.status = 'late' THEN 1 ELSE 0 END) as late_count
                FROM classes c
                LEFT JOIN campuses cp ON c.campus_id = cp.id
                LEFT JOIN students s ON c.id = s.class_id
                LEFT JOIN attendance a ON s.id = a.student_id AND a.date = ?
                GROUP BY c.id
            `;
            params = [selectedDate];
        } else {
            overviewQuery = `
                SELECT 
                    c.id, c.class_name, c.section,
                    COUNT(DISTINCT s.id) as total_students,
                    COUNT(DISTINCT a.id) as marked_count,
                    SUM(CASE WHEN a.status = 'present' THEN 1 ELSE 0 END) as present_count,
                    SUM(CASE WHEN a.status = 'absent' THEN 1 ELSE 0 END) as absent_count,
                    SUM(CASE WHEN a.status = 'late' THEN 1 ELSE 0 END) as late_count
                FROM classes c
                LEFT JOIN students s ON c.id = s.class_id AND s.campus_id = ?
                LEFT JOIN attendance a ON s.id = a.student_id AND a.date = ?
                WHERE c.campus_id = ?
                GROUP BY c.id
            `;
            params = [campusId, selectedDate, campusId];
        }

        console.log('[ATTENDANCE] Fetching overview for date:', selectedDate);
        console.log('[ATTENDANCE] Query params:', params);

        db.query(overviewQuery, params, (err, overview) => {
            if (err) {
                console.error('[ATTENDANCE ERROR] Overview query failed:', err);
                console.error('[ATTENDANCE ERROR] Query:', overviewQuery);
                console.error('[ATTENDANCE ERROR] Params:', params);
                // Return empty overview instead of crashing
                overview = [];
            }

            const classQuery = isSuperAdmin ? 'SELECT *, (SELECT campus_name FROM campuses WHERE id = classes.campus_id) as campus_name FROM classes' : 'SELECT * FROM classes WHERE campus_id = ?';
            const classParams = isSuperAdmin ? [] : [campusId];

            db.query(classQuery, classParams, (err2, classes) => {
                if (err2) {
                    console.error('[ATTENDANCE ERROR] Classes query failed:', err2);
                    classes = [];
                }

                console.log('[ATTENDANCE] Rendering overview with', overview.length, 'class stats');
                res.render('admin/attendance', {
                    classes,
                    overview: overview || [],
                    students: [],
                    selectedClass: null,
                    selectedDate,
                    viewMode: 'overview',
                    isSuperAdmin
                });
            });
        });
        return;
    }

    // Existing logic for specific class view
    const classQuery = isSuperAdmin ? 'SELECT * FROM classes' : 'SELECT * FROM classes WHERE campus_id = ?';
    const classParams = isSuperAdmin ? [] : [campusId];

    db.query(classQuery, classParams, (err, classes) => {
        const query = `
            SELECT s.id, u.full_name, s.roll_no, a.status 
            FROM students s
            JOIN users u ON s.user_id = u.id
            LEFT JOIN attendance a ON s.id = a.student_id AND a.date = ?
            WHERE s.class_id = ?
            ORDER BY s.roll_no
        `;
        db.query(query, [selectedDate, selectedClass], (err2, students) => {
            if (err2) console.error(err2);
            res.render('admin/attendance', {
                classes,
                students: students || [],
                selectedClass,
                selectedDate,
                overview: [],
                viewMode: 'detail',
                isSuperAdmin
            });
        });
    });
});

// 📊 EXPORT ATTENDANCE REPORT TO EXCEL
app.get('/admin/attendance/export/excel', async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const selectedDate = req.query.date || new Date().toISOString().split('T')[0];
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    try {
        let query = '';
        let params = [];

        if (isSuperAdmin) {
            query = `
                SELECT 
                    u.full_name as student_name, s.roll_no, c.class_name, c.section, 
                    cp.campus_name, a.status, a.date
                FROM attendance a
                JOIN students s ON a.student_id = s.id
                JOIN users u ON s.user_id = u.id
                JOIN classes c ON s.class_id = c.id
                LEFT JOIN campuses cp ON s.campus_id = cp.id
                WHERE a.date = ?
                ORDER BY cp.campus_name, c.class_name, s.roll_no
            `;
            params = [selectedDate];
        } else {
            query = `
                SELECT 
                    u.full_name as student_name, s.roll_no, c.class_name, c.section, 
                    a.status, a.date
                FROM attendance a
                JOIN students s ON a.student_id = s.id
                JOIN users u ON s.user_id = u.id
                JOIN classes c ON s.class_id = c.id
                WHERE a.date = ? AND s.campus_id = ?
                ORDER BY c.class_name, s.roll_no
            `;
            params = [selectedDate, campusId];
        }

        db.query(query, params, async (err, attendance) => {
            if (err) {
                console.error('[ATTENDANCE EXPORT ERROR]', err);
                console.error('[ATTENDANCE EXPORT ERROR] Query:', query);
                console.error('[ATTENDANCE EXPORT ERROR] Params:', params);
                return res.status(500).send('Error fetching attendance data: ' + err.message);
            }

            console.log(`[ATTENDANCE EXPORT] Found ${attendance.length} attendance records for ${selectedDate}`);

            // Create Excel Workbook
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Attendance Report');

            // Set column headers
            const columns = [
                { header: 'Date', key: 'date', width: 12 },
                { header: 'Roll No', key: 'roll_no', width: 10 },
                { header: 'Student Name', key: 'student_name', width: 25 },
                { header: 'Class', key: 'class_name', width: 15 },
                { header: 'Section', key: 'section', width: 12 },
                { header: 'Status', key: 'status', width: 12 }
            ];

            if (isSuperAdmin) {
                columns.splice(5, 0, { header: 'Campus', key: 'campus_name', width: 20 });
            }

            worksheet.columns = columns;

            // Style header row
            worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            worksheet.getRow(1).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF3B82F6' }
            };
            worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

            // Add data rows
            attendance.forEach(a => {
                const row = {
                    date: a.date instanceof Date ? a.date.toISOString().split('T')[0] : a.date,
                    roll_no: a.roll_no,
                    student_name: a.student_name,
                    class_name: a.class_name,
                    section: a.section,
                    status: (a.status || 'present').toUpperCase()
                };

                if (isSuperAdmin) {
                    row.campus_name = a.campus_name || '-';
                }

                worksheet.addRow(row);
            });

            // Set response headers
            const filename = `Attendance_Report_${selectedDate}.xlsx`;
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

            // Write to response
            await workbook.xlsx.write(res);
            res.end();

            console.log(`[EXPORT SUCCESS] Attendance report exported: ${filename}`);
        });
    } catch (error) {
        console.error('[EXPORT EXCEPTION]', error);
        res.status(500).send('Export failed');
    }
});


// Admin - Save Attendance
app.post('/admin/attendance/save', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { class_id, date, student_ids, status } = req.body;

    // We trust the class_id to imply campus, or fetch it if needed.
    // Ideally we should use the student's campus logic.
    const ids = Array.isArray(student_ids) ? student_ids : [student_ids];
    const statuses = Array.isArray(status) ? status : [status];

    if (!ids || ids.length === 0) return res.redirect('/admin/attendance');

    // Determine campus from first student or class (simplification for speed)
    // We'll stick to login campus ID for now, or fallback to class query if critical.
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const deleteQuery = 'DELETE FROM attendance WHERE date = ? AND student_id IN (SELECT id FROM students WHERE class_id = ?)';
    db.execute(deleteQuery, [date, class_id], (err) => {
        if (err) console.error(err);

        const insertQuery = 'INSERT INTO attendance (student_id, date, status, marked_by, campus_id) VALUES ?';
        const values = ids.map((id, index) => [id, date, statuses[index], req.session.user.id, campusId]);

        db.query(insertQuery, [values], (err2) => {
            if (err2) console.error(err2);
            res.redirect(`/admin/attendance?class_id=${class_id}&date=${date}`);
        });
    });
});


// Admin - Manage Notices
app.get('/admin/notices', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = isSuperAdmin ?
        'SELECT n.*, c.campus_name FROM notices n LEFT JOIN campuses c ON n.campus_id = c.id ORDER BY created_at DESC' :
        'SELECT * FROM notices WHERE campus_id = ? ORDER BY created_at DESC';
    const params = isSuperAdmin ? [] : [campusId];

    db.query(query, params, (err, notices) => {
        if (err) {
            console.error("Notices Query Error:", err);
            notices = [];
        }
        res.render('admin/notices', { notices: notices || [], isSuperAdmin });
    });
});

// Admin - Delete Notice - SECURED
app.get('/admin/notices/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = 'DELETE FROM notices WHERE id = ?';
    let params = [req.params.id];

    if (!isSuperAdmin) {
        query += ' AND campus_id = ?';
        params.push(campusId);
    }

    db.execute(query, params, (err) => {
        if (err) {
            console.error(err);
            return res.status(500).send('Error deleting notice');
        }
        res.redirect('/admin/notices');
    });
});

// 🔒 Multer for Notices with Security
const noticeStorage = multer.diskStorage({
    destination: (req, file, cb) => {
        const dir = './public/uploads/notices';
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        cb(null, dir);
    },
    filename: (req, file, cb) => {
        const sanitizedName = file.originalname.replace(/[^a-zA-Z0-9.-]/g, '_');
        cb(null, Date.now() + '-' + sanitizedName);
    }
});

const uploadNotice = multer({
    storage: noticeStorage,
    limits: {
        fileSize: 5 * 1024 * 1024 // 5MB for notices
    },
    fileFilter: (req, file, cb) => {
        const allowedMimes = [
            'application/pdf',
            'image/jpeg',
            'image/jpg',
            'image/png',
            'application/msword',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ];

        if (allowedMimes.includes(file.mimetype)) {
            cb(null, true);
        } else {
            cb(new Error('Invalid file type for notice. Only PDF, images, and documents allowed.'));
        }
    }
});

// Admin - Add Notice
// 🔒 CSRF Protected
app.post('/admin/notices/add', uploadNotice.single('attachment'), csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    // 🔒 SECURITY FIX: Sanitize inputs to prevent XSS
    const title = xss(req.body.title || 'Untitled Notice');
    const content = xss(req.body.content || '');

    // 🔒 Validate length
    if (title.length > 255) {
        return res.status(400).send('Title too long (max 255 characters)');
    }
    if (content.length > 10000) {
        return res.status(400).send('Content too long (max 10000 characters)');
    }

    const attachmentPath = req.file ? '/uploads/notices/' + req.file.filename : null;
    const authorId = req.session.user.id || null;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = 'INSERT INTO notices (title, content, posted_by, attachment_path, campus_id) VALUES (?, ?, ?, ?, ?)';

    // Explicitly handle parameters to avoid undefined
    db.execute(query, [title, content, authorId, attachmentPath, campusId], (err) => {
        if (err) {
            console.error(err);
            return res.status(500).send('Error posting notice: ' + err.message);
        }
        res.redirect('/admin/notices');
    });
});



// Admin - Library Management
app.get('/admin/library', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = isSuperAdmin ?
        'SELECT l.*, c.campus_name FROM library l LEFT JOIN campuses c ON l.campus_id = c.id ORDER BY uploaded_at DESC' :
        'SELECT * FROM library WHERE campus_id = ? ORDER BY uploaded_at DESC';
    const params = isSuperAdmin ? [] : [campusId];

    db.query(query, params, (err, books) => {
        if (err) {
            console.error("Library Query Error:", err);
            books = [];
        }
        res.render('admin/library', { books: books || [], isSuperAdmin });
    });
});

// Admin - Add Book Page (Simplified)
app.get('/admin/library/add', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    // Scoped for creating, but Super Admin can ideally pick campus. For now simple.
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    db.query('SELECT * FROM classes WHERE campus_id = ?', [campusId], (err, classes) => {
        if (err) console.error(err);
        res.render('admin/library_add', { classes: classes || [] });
    });
});

// Admin - Upload Book
app.post('/admin/library/add', upload.single('book_file'), csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { title, subject, target_grade } = req.body;
    const file = req.file;

    if (!file) return res.status(400).send('No file uploaded.');

    const file_path = '/uploads/library/' + file.filename;
    // Simple file type detection based on extension
    const ext = path.extname(file.originalname).toLowerCase();
    let file_type = 'other';
    if (['.pdf'].includes(ext)) file_type = 'pdf';
    else if (['.doc', '.docx'].includes(ext)) file_type = 'word';
    else if (['.jpg', '.jpeg', '.png'].includes(ext)) file_type = 'image';
    else if (['.mp4', '.avi', '.mov', '.mkv'].includes(ext)) file_type = 'video';
    else if (['.zip', '.rar'].includes(ext)) file_type = 'archive';

    const visibility = req.body.visibility || 'both';
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = 'INSERT INTO library (title, subject, target_grade, visibility, file_path, file_type, uploaded_by, campus_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?)';
    db.execute(query, [title, subject, target_grade || 'General', visibility, file_path, file_type, req.session.user.id, campusId], (err) => {
        if (err) {
            console.error(err);
            return res.status(500).send('Error uploading book');
        }
        res.redirect('/admin/library');
    });
});

// Admin - Delete Book - SECURED
app.get('/admin/library/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // First get path to delete file
    let validQuery = 'SELECT file_path FROM library WHERE id = ?';
    let validParams = [req.params.id];
    if (!isSuperAdmin) {
        validQuery += ' AND campus_id = ?';
        validParams.push(campusId);
    }

    db.query(validQuery, validParams, (err, results) => {
        if (results && results.length > 0) {
            const filePath = path.join(__dirname, 'public', results[0].file_path);
            if (fs.existsSync(filePath)) {
                fs.unlinkSync(filePath);
            }
            db.execute('DELETE FROM library WHERE id = ?', [req.params.id], (err2) => {
                if (err2) console.error(err2);
                res.redirect('/admin/library');
            });
        } else {
            res.redirect('/admin/library'); // Forbidden or not found
        }
    });
});

// Admin - Timetable Management
app.get('/admin/timetable', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // Get all classes for dropdown
    const classQuery = isSuperAdmin ? 'SELECT * FROM classes' : 'SELECT * FROM classes WHERE campus_id = ?';
    const classParams = isSuperAdmin ? [] : [campusId];

    db.query(classQuery, classParams, (err, classes) => {
        // Get all teachers for dropdown
        const teacherQuery = isSuperAdmin ?
            'SELECT t.id, u.full_name FROM teachers t JOIN users u ON t.user_id = u.id' :
            'SELECT t.id, u.full_name FROM teachers t JOIN users u ON t.user_id = u.id WHERE t.campus_id = ?';
        const teacherParams = isSuperAdmin ? [] : [campusId];

        db.query(teacherQuery, teacherParams, (err2, teachers) => {
            // Get timetable entries
            let timetableQuery = '';
            let ttParams = [];

            if (isSuperAdmin) {
                timetableQuery = `
                    SELECT t.*, c.class_name, c.section, u.full_name as teacher_name, cp.campus_name
                    FROM timetable t 
                    JOIN classes c ON t.class_id = c.id 
                    LEFT JOIN teachers tr ON t.teacher_id = tr.id 
                    LEFT JOIN users u ON tr.user_id = u.id 
                    LEFT JOIN campuses cp ON c.campus_id = cp.id
                    ORDER BY t.day, t.start_time
                 `;
            } else {
                timetableQuery = `
                    SELECT t.*, c.class_name, c.section, u.full_name as teacher_name 
                    FROM timetable t 
                    JOIN classes c ON t.class_id = c.id 
                    LEFT JOIN teachers tr ON t.teacher_id = tr.id 
                    LEFT JOIN users u ON tr.user_id = u.id 
                    WHERE c.campus_id = ?
                    ORDER BY t.day, t.start_time
                 `;
                ttParams = [campusId];
            }

            db.query(timetableQuery, ttParams, (err3, timetable) => {
                res.render('admin/timetable', { classes, teachers, timetable, isSuperAdmin });
            });
        });
    });
});

// Admin - Add Timetable Entry
app.post('/admin/timetable/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { class_id, subject, day, start_time, end_time, teacher_id } = req.body;
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    const query = 'INSERT INTO timetable (class_id, subject, day, start_time, end_time, teacher_id, campus_id) VALUES (?, ?, ?, ?, ?, ?, ?)';
    db.execute(query, [class_id, subject, day, start_time, end_time, teacher_id || null, campusId], (err) => {
        if (err) {
            console.error("[TIMETABLE ERROR]", err);
            return res.status(500).send("Error saving schedule: " + err.message);
        }
        res.redirect('/admin/timetable?success=Scheduled+Successfully');
    });
});

// Admin - Delete Timetable Entry - SECURED
app.get('/admin/timetable/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = 'DELETE FROM timetable WHERE id = ?';
    let params = [req.params.id];

    if (!isSuperAdmin) {
        query += ' AND campus_id = ?';
        params.push(campusId);
    }

    db.execute(query, params, (err) => {
        if (err) console.error(err);
        res.redirect('/admin/timetable');
    });
});


// 🏢 CAMPUS MANAGEMENT ROUTES
// ==========================================

// List Campuses
app.get('/admin/campuses', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const isSuperAdmin = (req.session.user.username === 'admin' && Number(req.session.user.campus_id) === 1);

    db.query("SELECT * FROM campuses ORDER BY id ASC", (err, campuses) => {
        if (err) { console.error(err); return res.send("Error"); }
        res.render('admin/campuses', { campuses, user: req.session.user, isSuperAdmin });
    });
});

// Add Campus
// Add Campus & Auto-create Admin
app.post('/admin/campuses/add', csrfProtection, async (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    // 🔒 SUPER ADMIN ONLY CHECK
    const isSuperAdmin = (req.session.user.username === 'admin' && Number(req.session.user.campus_id) === 1);
    if (!isSuperAdmin) {
        return res.status(403).send("Unauthorized: Only Super Admin can add campuses.");
    }

    let { campus_name, campus_code, city, address } = req.body;
    campus_code = campus_code.trim().toUpperCase();

    // 1. Create Campus
    db.query("INSERT INTO campuses (campus_name, campus_code, city, address) VALUES (?, ?, ?, ?)",
        [campus_name, campus_code, city, address],
        async (err, result) => {
            if (err) {
                console.error("Add Campus Error:", err);
                return res.redirect('/admin/campuses?error=create_failed');
            }

            const newCampusId = result.insertId;

            // 2. Create Default Admin for this Campus
            try {
                // Use provided password or default to 'admin123'
                const rawPassword = req.body.admin_password || 'admin123';
                const hashedPassword = await bcrypt.hash(rawPassword, 10);

                // We use 'admin' as username, which is now allowed per campus
                db.execute(
                    'INSERT INTO users (username, password, role, full_name, campus_id, email) VALUES (?, ?, "admin", ?, ?, ?)',
                    ['admin', hashedPassword, `Admin - ${campus_name}`, newCampusId, `admin@${campus_code.toLowerCase()}.com`],
                    (err2) => {
                        if (err2) console.error("Auto-Admin Creation Failed:", err2);
                        else console.log(`Created default admin for Campus ${newCampusId} with password length ${rawPassword.length}`);

                        res.redirect('/admin/campuses?success=created_with_admin');
                    }
                );
            } catch (e) {
                console.error(e);
                res.redirect('/admin/campuses');
            }
        });
});

// Delete Campus (Protection for Main Campus)
app.get('/admin/campuses/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    // 🔒 SUPER ADMIN ONLY CHECK
    const isSuperAdmin = (req.session.user.username === 'admin' && Number(req.session.user.campus_id) === 1);
    if (!isSuperAdmin) {
        return res.status(403).send("Unauthorized: Only Super Admin can delete campuses.");
    }

    const campusId = parseInt(req.params.id);
    if (campusId === 1) return res.send("Cannot delete Main Campus!");

    db.query("DELETE FROM campuses WHERE id = ?", [campusId], (err) => {
        if (err) console.error(err);
        res.redirect('/admin/campuses');
    });
});
// Admin - Security Audit Logs
app.get('/admin/audit-logs', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const limit = 500; // Increased to 500 for better visibility
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);


    // 1. Get Total Count (for Stats)
    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);

    let countQuery = "SELECT COUNT(*) as total FROM audit_logs";
    let countParams = [];
    if (!isSuperAdmin) {
        countQuery += " WHERE campus_id = ?";
        countParams.push(campusId);
    }

    db.query(countQuery, countParams, (err1, countRes) => {
        const totalCount = countRes ? countRes[0].total : 0;

        // 2. Get Recent Logs (for Table)
        let sql = `
            SELECT l.*, u.username as user_identifier, cp.campus_name 
            FROM audit_logs l 
            LEFT JOIN users u ON l.user_id = u.id 
            LEFT JOIN campuses cp ON l.campus_id = cp.id
        `;
        let logParams = [];

        if (!isSuperAdmin) {
            sql += " WHERE l.campus_id = ?";
            logParams.push(campusId);
        }

        sql += " ORDER BY l.created_at DESC LIMIT ?";
        logParams.push(limit);

        db.query(sql, logParams, (err, logs) => {
            if (err) {
                console.error(err);
                return res.send("Error fetching logs: " + err.message);
            }
            res.render('admin/audit_logs', { logs, totalCount, isSuperAdmin });
        });
    });
});

// Admin - Payroll Management
app.get('/admin/payroll', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const selectedMonth = req.query.month || new Date().toLocaleString('default', { month: 'long' });
    const selectedYear = req.query.year || new Date().getFullYear();

    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // Fetch teachers with their payroll record for selected month/year (Scoped to Campus)
    const query = `
        SELECT t.id as teacher_id, u.full_name, t.salary, p.id as payroll_id, p.status, p.paid_date
        FROM teachers t
        JOIN users u ON t.user_id = u.id
        LEFT JOIN payroll p ON t.id = p.teacher_id AND p.month = ? AND p.year = ?
        WHERE t.campus_id = ?
    `;

    db.query(query, [selectedMonth, selectedYear, campusId], (err, results) => {
        if (err) console.error(err);
        res.render('admin/payroll', {
            results,
            selectedMonth,
            selectedYear,
            months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        });
    });
});

// Admin - Mark Salary as Paid - SECURED
// 🔒 CSRF Protected
app.post('/admin/payroll/pay', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.status(403).send("Forbidden");
    const { teacher_id, amount, month, year } = req.body;

    // DEBUG LOGGING
    console.log('=== PAYROLL PAYMENT REQUEST ===');
    console.log('Received Data:', { teacher_id, amount, month, year });
    console.log('User:', req.session.user.username, 'Campus:', req.session.campus);

    const isSuperAdmin = (req.session.user.username === 'admin' && req.session.user.campus_id === 1);
    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    // Security Check: Ensure teacher matches campus
    const secureTeacherQuery = isSuperAdmin ?
        'SELECT id, campus_id FROM teachers WHERE id = ?' :
        'SELECT id, campus_id FROM teachers WHERE id = ? AND campus_id = ?';
    const secureParams = isSuperAdmin ? [teacher_id] : [teacher_id, campusId];

    console.log('Teacher Query:', secureTeacherQuery, 'Params:', secureParams);

    db.query(secureTeacherQuery, secureParams, (errT, teacherRes) => {
        if (errT) {
            console.error('Teacher Query Error:', errT);
            return res.send("Error: Database error - " + errT.message);
        }

        if (teacherRes.length === 0) {
            console.error('Teacher not found. ID:', teacher_id, 'Campus:', campusId);
            return res.send("Error: Teacher not found or access denied.");
        }

        const authorizedTeacher = teacherRes[0];
        console.log('Teacher Found:', authorizedTeacher);

        // Use teacher's campus for the record (important for Super Admin paying for a specific campus teacher)
        const recordCampusId = authorizedTeacher.campus_id;

        const checkQuery = "SELECT id FROM payroll WHERE teacher_id = ? AND month = ? AND year = ?";
        console.log('Check Query:', checkQuery, 'Params:', [teacher_id, month, year]);

        db.query(checkQuery, [teacher_id, month, year], (err, exists) => {
            if (err) {
                console.error('Check Query Error:', err);
                return res.send("Error: " + err.message);
            }

            console.log('Existing Records Found:', exists.length, exists);

            if (exists.length > 0) {
                console.log('Updating existing record ID:', exists[0].id);
                db.execute("UPDATE payroll SET status = 'paid', paid_date = CURRENT_DATE WHERE id = ?", [exists[0].id], (err2) => {
                    if (err2) {
                        console.error('Update Error:', err2);
                        return res.send("Error updating: " + err2.message);
                    }
                    console.log('✓ Successfully updated payroll record');
                    res.redirect(`/admin/payroll?month=${month}&year=${year}`);
                });
            } else {
                console.log('Creating new payroll record');
                db.execute("INSERT INTO payroll (teacher_id, amount, month, year, status, paid_date, campus_id) VALUES (?, ?, ?, ?, 'paid', CURRENT_DATE, ?)",
                    [teacher_id, amount, month, year, recordCampusId], (err2) => {
                        if (err2) {
                            console.error('Insert Error:', err2);
                            return res.send("Error inserting: " + err2.message);
                        }
                        console.log('✓ Successfully created payroll record');
                        res.redirect(`/admin/payroll?month=${month}&year=${year}`);
                    });
            }
        });
    });
});


app.get('/teacher/dashboard', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    const teacherQuery = 'SELECT * FROM teachers WHERE user_id = ?';
    db.query(teacherQuery, [req.session.user.id], (err, teacherResults) => {
        db.query('SELECT * FROM notices ORDER BY created_at DESC LIMIT 5', (err2, notices) => {
            res.render('teacher/dashboard', {
                teacherData: teacherResults[0] || null,
                notices: notices
            });
        });
    });
});

// Teacher - Attendance View
app.get('/teacher/attendance', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    const selectedClass = req.query.class_id || null;
    const selectedDate = req.query.date || new Date().toISOString().split('T')[0];

    // 1. Get Teacher Profile
    db.query('SELECT * FROM teachers WHERE user_id = ?', [req.session.user.id], (err, teacherResult) => {
        if (err || teacherResult.length === 0) return res.send("Teacher Profile Not Found.");

        const teacher = teacherResult[0];

        // 2. Fetch ONLY classes where this teacher is the CLASS TEACHER
        db.query('SELECT * FROM classes WHERE class_teacher_id = ?', [teacher.id], (err2, classes) => {
            if (err2) return res.send("Error fetching classes.");

            if (selectedClass) {
                // Security: Ensure the teacher is actually allowed to view this class
                const isAssigned = classes.find(c => c.id == selectedClass);
                if (!isAssigned) return res.send("Access Denied: You are not the Class Teacher for this class.");

                const query = `
                    SELECT s.id, u.full_name, s.roll_no, a.status 
                    FROM students s
                    JOIN users u ON s.user_id = u.id
                    LEFT JOIN attendance a ON s.id = a.student_id AND a.date = ?
                    WHERE s.class_id = ?
                    ORDER BY s.roll_no
                `;
                db.query(query, [selectedDate, selectedClass], (err3, students) => {
                    res.render('teacher/attendance', {
                        teacherData: teacher,
                        classes: classes || [],
                        students: students || [],
                        selectedClass,
                        selectedDate
                    });
                });
            } else {
                res.render('teacher/attendance', {
                    teacherData: teacher,
                    classes: classes || [],
                    students: [],
                    selectedClass: null,
                    selectedDate
                });
            }
        });
    });
});

// Teacher - Save Attendance
app.post('/teacher/attendance/save', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');
    const { class_id, date, student_ids, status } = req.body;

    const ids = Array.isArray(student_ids) ? student_ids : [student_ids];
    const statuses = Array.isArray(status) ? status : [status];

    const deleteQuery = 'DELETE FROM attendance WHERE date = ? AND student_id IN (SELECT id FROM students WHERE class_id = ?)';
    db.execute(deleteQuery, [date, class_id], (err) => {
        const insertQuery = 'INSERT INTO attendance (student_id, date, status, marked_by) VALUES ?';
        const values = ids.map((id, index) => [id, date, statuses[index], req.session.user.id]);
        db.query(insertQuery, [values], (err2) => {
            res.redirect(`/teacher/attendance?class_id=${class_id}&date=${date}`);
        });
    });
});

// Teacher - Exam Results
app.get('/teacher/exams', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    const selectedExam = req.query.exam_id || null;
    const selectedClass = req.query.class_id || null;

    db.query('SELECT * FROM teachers WHERE user_id = ?', [req.session.user.id], (err, teacherResult) => {
        if (err || teacherResult.length === 0) return res.send("Teacher profile not found.");
        const teacher = teacherResult[0];

        db.query('SELECT * FROM exams ORDER BY id DESC', (err2, exams) => {
            // 🛡️ SECURITY: Only show classes where this teacher is the Class Teacher
            db.query('SELECT * FROM classes WHERE class_teacher_id = ?', [teacher.id], (err3, classes) => {
                if (selectedExam && selectedClass) {
                    // double check the teacher is the class teacher for the selected class
                    const isAssigned = classes.find(c => c.id == selectedClass);
                    if (!isAssigned) return res.send("Access Denied: You are not the Class Teacher for this class.");

                    const query = `
                        SELECT s.id, u.full_name, s.roll_no, er.marks_obtained 
                        FROM students s
                        JOIN users u ON s.user_id = u.id
                        LEFT JOIN exam_results er ON s.id = er.student_id AND er.exam_id = ? AND er.subject = ?
                        WHERE s.class_id = ?
                        ORDER BY s.roll_no
                    `;
                    db.query(query, [selectedExam, teacher.subject, selectedClass], (err4, students) => {
                        res.render('teacher/exams', {
                            teacherData: teacher,
                            exams: exams || [],
                            classes: classes || [],
                            students: students || [],
                            selectedExam,
                            selectedClass
                        });
                    });
                } else {
                    res.render('teacher/exams', {
                        teacherData: teacher,
                        exams: exams || [],
                        classes: classes || [],
                        students: [],
                        selectedExam: null,
                        selectedClass: null
                    });
                }
            });
        });
    });
});

// Teacher - Save Exam Results
app.post('/teacher/exams/save', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');
    const { exam_id, class_id, subject, total_marks, student_ids, marks } = req.body;

    // 🛡️ SECURITY CHECK: Verify if the teacher is the Class Teacher
    db.query('SELECT id FROM teachers WHERE user_id = ?', [req.session.user.id], (err, teacherResult) => {
        if (err || teacherResult.length === 0) return res.send("Teacher not found.");
        const teacherId = teacherResult[0].id;

        db.query('SELECT id FROM classes WHERE id = ? AND class_teacher_id = ?', [class_id, teacherId], (errCheck, checkRes) => {
            if (errCheck || checkRes.length === 0) return res.status(403).send("Unauthorized: You can only save results for your own class.");

            const ids = Array.isArray(student_ids) ? student_ids : [student_ids];
            const marksList = Array.isArray(marks) ? marks : [marks];

            const deleteQuery = 'DELETE FROM exam_results WHERE exam_id = ? AND subject = ? AND student_id IN (SELECT id FROM students WHERE class_id = ?)';
            db.execute(deleteQuery, [exam_id, subject, class_id], (err) => {
                const insertQuery = 'INSERT INTO exam_results (exam_id, student_id, subject, marks_obtained, total_marks, grade) VALUES ?';
                const values = ids.map((id, index) => {
                    const m = parseFloat(marksList[index]) || 0;
                    const percentage = (m / total_marks) * 100;
                    let grade = 'F';
                    if (percentage >= 80) grade = 'A+';
                    else if (percentage >= 70) grade = 'A';
                    else if (percentage >= 60) grade = 'B';
                    else if (percentage >= 50) grade = 'C';
                    else if (percentage >= 40) grade = 'D';

                    return [exam_id, id, subject, m, total_marks, grade];
                });

                db.query(insertQuery, [values], (err2) => {
                    res.redirect(`/teacher/exams?exam_id=${exam_id}&class_id=${class_id}`);
                });
            });
        });
    });
});

// 📚 TEACHER HOMEWORK MANAGEMENT
// ==========================================

// 1. List Homework
app.get('/teacher/homework', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    db.query('SELECT * FROM teachers WHERE user_id = ?', [req.session.user.id], (err, teacherResult) => {
        if (err || teacherResult.length === 0) return res.send("Teacher not found.");
        const teacher = teacherResult[0];

        // Fetch all homework created by this teacher
        const hwQuery = `
            SELECT h.*, c.class_name, c.section,
            (SELECT COUNT(*) FROM homework_submissions WHERE homework_id = h.id) as total_submissions
            FROM homework h
            JOIN classes c ON h.class_id = c.id
            WHERE h.teacher_id = ?
            ORDER BY h.created_at DESC
        `;

        db.query(hwQuery, [teacher.id], (err2, homework) => {
            // Fetch relevant classes for the 'Add Homework' modal
            const classQuery = `
                SELECT DISTINCT c.id, c.class_name, c.section
                FROM classes c
                WHERE c.class_teacher_id = ?
                OR c.id IN (SELECT class_id FROM timetable WHERE teacher_id = ?)
            `;

            db.query(classQuery, [teacher.id, teacher.id], (err3, classes) => {
                res.render('teacher/homework', {
                    teacherData: teacher,
                    homework: homework || [],
                    classes: classes || [],
                    error: req.query.error || null,
                    success: req.query.success || null
                });
            });
        });
    });
});

// 2. Create Homework
app.post('/teacher/homework/add', uploadHomework.single('attachment'), csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');
    const { class_id, subject, title, description, due_date, total_marks } = req.body;
    const attachment_path = req.file ? `/uploads/homework/${req.file.filename}` : null;

    db.query('SELECT id, campus_id FROM teachers WHERE user_id = ?', [req.session.user.id], (err, teacherResult) => {
        if (err || teacherResult.length === 0) return res.send("Unauthorized");
        const teacher = teacherResult[0];

        const query = 'INSERT INTO homework (campus_id, teacher_id, class_id, subject, title, description, due_date, total_marks, attachment_path) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)';
        db.execute(query, [teacher.campus_id, teacher.id, class_id, subject, title, description, due_date, total_marks, attachment_path], (err2) => {
            if (err2) {
                console.error(err2);
                return res.redirect('/teacher/homework?error=Failed to create homework');
            }
            res.redirect('/teacher/homework?success=Homework created successfully');
        });
    });
});

// 3. Delete Homework
app.get('/teacher/homework/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    db.query('SELECT id FROM teachers WHERE user_id = ?', [req.session.user.id], (err, teacherResult) => {
        const teacherId = teacherResult[0].id;
        db.execute('DELETE FROM homework WHERE id = ? AND teacher_id = ?', [req.params.id, teacherId], (err2) => {
            res.redirect('/teacher/homework?success=Homework deleted');
        });
    });
});

// 4. View Submissions
app.get('/teacher/homework/submissions/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    const homeworkId = req.params.id;

    db.query('SELECT * FROM teachers WHERE user_id = ?', [req.session.user.id], (errT, teacherResult) => {
        if (errT || teacherResult.length === 0) return res.send("Teacher profile not found.");
        const teacher = teacherResult[0];

        db.query('SELECT h.*, c.class_name, c.section FROM homework h JOIN classes c ON h.class_id = c.id WHERE h.id = ?', [homeworkId], (err, hwResult) => {
            if (err || hwResult.length === 0) return res.redirect('/teacher/homework');
            const homework = hwResult[0];

            // Fetch all students in this class and their submissions
            const submissionQuery = `
                SELECT s.id as student_id, u.full_name, s.roll_no, 
                       sub.id as submission_id, sub.submission_file, sub.submitted_at, sub.status, sub.marks_obtained, sub.teacher_remarks
                FROM students s
                JOIN users u ON s.user_id = u.id
                LEFT JOIN homework_submissions sub ON s.id = sub.student_id AND sub.homework_id = ?
                WHERE s.class_id = ?
                ORDER BY s.roll_no
            `;

            db.query(submissionQuery, [homeworkId, homework.class_id], (err2, submissions) => {
                res.render('teacher/homework_submissions', {
                    teacherData: teacher,
                    homework,
                    submissions: submissions || []
                });
            });
        });
    });
});

// 5. Grade Homework
app.post('/teacher/homework/grade', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');
    const { submission_id, marks, remarks } = req.body;

    db.execute(
        'UPDATE homework_submissions SET marks_obtained = ?, teacher_remarks = ?, graded_at = NOW() WHERE id = ?',
        [marks, remarks, submission_id],
        (err) => {
            if (err) console.error(err);
            res.json({ success: true });
        }
    );
});

// Teacher - Timetable
app.get('/teacher/timetable', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    const teacherQuery = 'SELECT id, subject FROM teachers WHERE user_id = ?';
    db.query(teacherQuery, [req.session.user.id], (err, teacher) => {
        const query = `
            SELECT t.*, c.class_name, c.section 
            FROM timetable t
            JOIN classes c ON t.class_id = c.id
            WHERE t.teacher_id = ?
            ORDER BY FIELD(day, 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'), start_time
        `;
        db.query(query, [teacher[0].id], (err2, timetable) => {
            res.render('teacher/timetable', {
                teacherData: teacher[0],
                timetable
            });
        });
    });
});


// Teacher - My Students (Block System)
app.get('/teacher/students', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'teacher') return res.redirect('/login');

    const teacherQuery = 'SELECT id, subject FROM teachers WHERE user_id = ?';
    db.query(teacherQuery, [req.session.user.id], (err, teacher) => {
        if (err || !teacher || teacher.length === 0) return res.redirect('/login');

        const classQuery = 'SELECT DISTINCT c.id, c.class_name, c.section FROM classes c JOIN timetable t ON c.id = t.class_id WHERE t.teacher_id = ?';
        db.query(classQuery, [teacher[0].id], (err2, classes) => {
            const selectedClass = req.query.class_id || null;

            if (selectedClass) {
                // Security Check: Ensure teacher is assigned to this class
                const isAssigned = classes.some(c => c.id == selectedClass);
                if (!isAssigned && classes.length > 0) {
                    return res.redirect('/teacher/students');
                }

                const studentsQuery = `
                    SELECT s.*, u.full_name, u.email 
                    FROM students s
                    JOIN users u ON s.user_id = u.id
                    WHERE s.class_id = ?
                    ORDER BY s.roll_no
                `;
                db.query(studentsQuery, [selectedClass], (err3, students) => {
                    const currentClass = classes.find(c => c.id == selectedClass);
                    res.render('teacher/students', {
                        teacherData: teacher[0],
                        classes,
                        students,
                        selectedClass,
                        currentClass,
                        viewMode: 'students'
                    });
                });
            } else {
                res.render('teacher/students', {
                    teacherData: teacher[0],
                    classes,
                    students: [],
                    selectedClass: null,
                    viewMode: 'blocks'
                });
            }
        });
    });
});

app.get('/student/dashboard', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');

    const studentQuery = `
        SELECT s.*, c.class_name, c.section 
        FROM students s 
        JOIN classes c ON s.class_id = c.id 
        WHERE s.user_id = ?
    `;
    db.query(studentQuery, [req.session.user.id], (err, studentResults) => {
        const studentData = studentResults[0] || null;
        if (!studentData) return res.send('Student profile not found');

        // Fetch Attendance Stats
        const attQuery = "SELECT COUNT(*) as total, SUM(CASE WHEN status = 'present' THEN 1 ELSE 0 END) as present FROM attendance WHERE student_id = ?";
        db.query(attQuery, [studentData.id], (errAtt, attStats) => {
            const attPercent = attStats[0].total > 0 ? Math.round((attStats[0].present / attStats[0].total) * 100) : 0;

            // Fetch Latest Exam Result
            const resQuery = "SELECT * FROM exam_results WHERE student_id = ? ORDER BY id DESC LIMIT 5";
            db.query(resQuery, [studentData.id], (errRes, results) => {

                db.query('SELECT * FROM vouchers WHERE student_id = ? ORDER BY academic_year DESC LIMIT 1', [studentData.id], (err2, vouchers) => {
                    db.query('SELECT * FROM notices ORDER BY created_at DESC LIMIT 5', (err3, notices) => {
                        // Fetch Class Fellows
                        const fellowQuery = `
                            SELECT u.full_name, s.roll_no 
                            FROM students s 
                            JOIN users u ON s.user_id = u.id 
                            WHERE s.class_id = ? AND s.id != ? 
                            LIMIT 6
                        `;
                        db.query(fellowQuery, [studentData.class_id, studentData.id], (err4, fellows) => {
                            res.render('student/dashboard', {
                                studentData,
                                vouchers,
                                notices,
                                attPercent,
                                results,
                                fellows: fellows || []
                            });
                        });
                    });
                });
            });
        });
    });
});

// Student - Results
app.get('/student/results', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');

    db.query('SELECT id FROM students WHERE user_id = ?', [req.session.user.id], (err, std) => {
        const query = `
            SELECT er.*, e.exam_name 
            FROM exam_results er
            JOIN exams e ON er.exam_id = e.id
            WHERE er.student_id = ?
            ORDER BY e.id DESC
        `;
        db.query(query, [std[0].id], (err2, results) => {
            // Group results by exam
            const groupedResults = results.reduce((acc, current) => {
                if (!acc[current.exam_name]) acc[current.exam_name] = [];
                acc[current.exam_name].push(current);
                return acc;
            }, {});

            res.render('student/results', { groupedResults });
        });
    });
});

// Student - Fees (Voucher Ledger View)
app.get('/student/fees', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');

    db.query('SELECT id FROM students WHERE user_id = ?', [req.session.user.id], (err, std) => {
        // Handle DB error or student record not found
        if (err) {
            console.error('Student Fees - DB Error fetching student:', err);
            return res.render('student/fees', { vouchers: [], items: [] });
        }
        if (!std || std.length === 0) {
            console.warn('Student Fees - No student record found for user_id:', req.session.user.id);
            return res.render('student/fees', { vouchers: [], items: [] });
        }

        const studentId = std[0].id;

        const query = `
            SELECT v.*, u.full_name, c.class_name, c.section, s.roll_no
            FROM vouchers v
            JOIN students s ON v.student_id = s.id
            JOIN users u ON s.user_id = u.id
            JOIN classes c ON s.class_id = c.id
            WHERE v.student_id = ?
            ORDER BY v.academic_year DESC
        `;

        db.query(query, [studentId], (err2, vouchers) => {
            if (err2) {
                console.error('Student Fees - Voucher Query Error:', err2);
                return res.render('student/fees', { vouchers: [], items: [] });
            }

            // Also fetch ALL monthly fee items for all vouchers of this student
            if (vouchers && vouchers.length > 0) {
                const voucherIds = vouchers.map(v => v.id);
                const placeholders = voucherIds.map(() => '?').join(',');
                db.query(
                    `SELECT * FROM fees WHERE voucher_id IN (${placeholders}) ORDER BY id DESC`,
                    voucherIds,
                    (err3, items) => {
                        if (err3) {
                            console.error('Student Fees - Items Query Error:', err3);
                            return res.render('student/fees', { vouchers, items: [] });
                        }
                        res.render('student/fees', { vouchers, items: items || [] });
                    }
                );
            } else {
                res.render('student/fees', { vouchers: [], items: [] });
            }
        });
    });
});


// Student - Timetable
app.get('/student/timetable', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');
    db.query('SELECT class_id FROM students WHERE user_id = ?', [req.session.user.id], (err, std) => {
        const query = `
            SELECT t.*, u.full_name as teacher_member 
            FROM timetable t
            LEFT JOIN teachers tr ON t.teacher_id = tr.id
            LEFT JOIN users u ON tr.user_id = u.id
            WHERE t.class_id = ?
            ORDER BY FIELD(day, 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'), start_time
        `;
        db.query(query, [std[0].class_id], (err2, timetable) => {
            res.render('student/timetable', { timetable });
        });
    });
});

// Student - Attendance
app.get('/student/attendance', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');
    db.query('SELECT id FROM students WHERE user_id = ?', [req.session.user.id], (err, std) => {
        const query = "SELECT * FROM attendance WHERE student_id = ? ORDER BY date DESC";
        db.query(query, [std[0].id], (err2, attendanceLog) => {
            const stats = attendanceLog.reduce((acc, curr) => {
                acc[curr.status]++;
                acc.total++;
                return acc;
            }, { present: 0, absent: 0, late: 0, leave: 0, total: 0 });

            res.render('student/attendance', { attendanceLog, stats });
        });
    });
});

// 🎒 STUDENT HOMEWORK
// ==========================================

// 1. List Homework for Student
app.get('/student/homework', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');

    db.query('SELECT s.*, c.class_name, c.section FROM students s JOIN classes c ON s.class_id = c.id WHERE s.user_id = ?', [req.session.user.id], (err, stdResults) => {
        if (err || stdResults.length === 0) return res.send("Student profile not found.");
        const student = stdResults[0];

        // Fetch all homework for this student's class, including their own submission status
        const hwQuery = `
            SELECT h.*, u.full_name as teacher_name,
                   sub.id as submission_id, sub.submitted_at, sub.status as submission_status, sub.marks_obtained, sub.teacher_remarks
            FROM homework h
            JOIN teachers t ON h.teacher_id = t.id
            JOIN users u ON t.user_id = u.id
            LEFT JOIN homework_submissions sub ON h.id = sub.homework_id AND sub.student_id = ?
            WHERE h.class_id = ?
            ORDER BY h.due_date ASC
        `;

        db.query(hwQuery, [student.id, student.class_id], (err2, homework) => {
            res.render('student/homework', {
                studentData: student,
                homework: homework || [],
                success: req.query.success || null,
                error: req.query.error || null
            });
        });
    });
});

// 2. Submit Homework
app.post('/student/homework/submit', uploadSubmission.single('submission'), csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'student') return res.redirect('/login');
    const { homework_id } = req.body;
    const submission_file = req.file ? `/uploads/submissions/${req.file.filename}` : null;

    if (!submission_file) return res.redirect('/student/homework?error=Please select a file to upload.');

    db.query('SELECT id, class_id, campus_id FROM students WHERE user_id = ?', [req.session.user.id], (err, stdResults) => {
        const student = stdResults[0];

        // Fetch due date to check for lateness
        db.query('SELECT due_date FROM homework WHERE id = ?', [homework_id], (err2, hw) => {
            const dueDate = new Date(hw[0].due_date);
            const now = new Date();
            const status = now > dueDate ? 'late' : 'on_time';

            const query = 'INSERT INTO homework_submissions (campus_id, homework_id, student_id, submission_file, status) VALUES (?, ?, ?, ?, ?)';
            db.execute(query, [student.campus_id, homework_id, student.id, submission_file, status], (err3) => {
                if (err3) {
                    console.error(err3);
                    return res.redirect('/student/homework?error=Submission failed.');
                }
                res.redirect('/student/homework?success=Homework submitted successfully.');
            });
        });
    });
});


// Parent Dashboard
app.get('/parent/dashboard', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'parent') return res.redirect('/login');

    // Find child linked to this parent
    const childQuery = `
        SELECT s.*, u.full_name, c.class_name, c.section 
        FROM students s 
        JOIN users u ON s.user_id = u.id
        JOIN classes c ON s.class_id = c.id 
        WHERE s.parent_id = ?
    `;
    db.query(childQuery, [req.session.user.id], (err, childResults) => {
        const childData = childResults[0] || null;
        if (childData) {
            db.query('SELECT * FROM vouchers WHERE student_id = ? ORDER BY academic_year DESC', [childData.id], (err2, vouchers) => {
                res.render('parent/dashboard', { childData, vouchers });
            });
        } else {
            res.render('parent/dashboard', { childData: null, vouchers: [] });
        }
    });
});



// Public/Student View Library (Read-only) - Can be accessed by specific roles
app.get('/library', (req, res) => {
    if (!req.session.user) return res.redirect('/login');

    // Trim and lowercase the role for robust comparison
    const role = String(req.session.user.role).toLowerCase().trim();
    const userId = req.session.user.id;

    console.log(`[Library] Access by: ${role} (UID: ${userId})`);

    if (role === 'student') {
        const studentQuery = `
            SELECT c.class_name, c.section 
            FROM students s 
            JOIN classes c ON s.class_id = c.id 
            WHERE s.user_id = ?
        `;
        db.query(studentQuery, [userId], (err, studentResults) => {
            if (err) {
                console.error("[Library] DB Error:", err);
                return res.status(500).send("An error occurred");
            }

            if (studentResults.length === 0) {
                console.warn(`[Library] Student UID ${userId} not found in 'students' table. Showing General only.`);
                db.query("SELECT * FROM library WHERE target_grade = 'General' AND (visibility = 'student' OR visibility = 'both') ORDER BY uploaded_at DESC", (err2, books) => {
                    if (err2) {
                        console.error("[Library] Query Error:", err2);
                        return res.status(500).send("Database Error: " + err2.message);
                    }
                    res.render('library_view', { books: books || [], role: 'student', studentClass: 'No Class Assigned' });
                });
            } else {
                const studentClass = `${studentResults[0].class_name} - ${studentResults[0].section}`;
                console.log(`[Library] Student Class: "${studentClass}"`);

                const libraryQuery = "SELECT * FROM library WHERE (target_grade = 'General' OR target_grade = ?) AND (visibility = 'student' OR visibility = 'both') ORDER BY uploaded_at DESC";
                db.query(libraryQuery, [studentClass], (err2, books) => {
                    if (err2) {
                        console.error("[Library] Query Error:", err2);
                        return res.status(500).send("Database Error: " + err2.message);
                    }
                    const foundGrades = [...new Set((books || []).map(b => b.target_grade))];
                    console.log(`[Library] Query returned ${(books || []).length} books. Grades found: ${foundGrades.join(', ')}`);
                    res.render('library_view', { books: books || [], role: 'student', studentClass: studentClass });
                });
            }
        });
    } else if (role === 'teacher') {
        // Teachers see everything targeted for 'teacher' or 'both'
        // They can also see any class material if Admin marked it for teachers
        db.query("SELECT * FROM library WHERE visibility = 'teacher' OR visibility = 'both' ORDER BY uploaded_at DESC", (err, books) => {
            if (err) console.error(err);
            res.render('library_view', { books: books || [], role });
        });
    } else if (role === 'admin') {
        // Admin sees EVERYTHING
        db.query('SELECT * FROM library ORDER BY uploaded_at DESC', (err, books) => {
            if (err) console.error(err);
            res.render('library_view', { books: books || [], role });
        });
    } else {
        // Parents/Others only see General for Students
        db.query("SELECT * FROM library WHERE target_grade = 'General' AND (visibility = 'student' OR visibility = 'both') ORDER BY uploaded_at DESC", (err, books) => {
            if (err) console.error(err);
            res.render('library_view', { books: books || [], role });
        });
    }
});

// Admin - Email Communication View
app.get('/admin/email', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    res.render('admin/email');
});

// Admin - Send Email Logic
app.post('/admin/send-email', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { recipient_group, subject, message } = req.body;

    const campusId = req.session.campus ? req.session.campus.id : (req.session.user ? req.session.user.campus_id : 1);

    let query = '';
    let params = [];
    if (recipient_group === 'students') {
        query = "SELECT email FROM users WHERE role = 'student' AND email IS NOT NULL AND email != '' AND campus_id = ?";
        params = [campusId];
    } else if (recipient_group === 'teachers') {
        query = "SELECT email FROM users WHERE role = 'teacher' AND email IS NOT NULL AND email != '' AND campus_id = ?";
        params = [campusId];
    } else {
        query = "SELECT email FROM users WHERE (role = 'student' OR role = 'teacher') AND email IS NOT NULL AND email != '' AND campus_id = ?";
        params = [campusId];
    }

    db.query(query, params, async (err, results) => {
        if (err) {
            console.error(err);
            return res.render('admin/email', { error: 'Database error fetching recipients.' });
        }

        const emails = results.map(r => r.email).filter(e => e && e.includes('@'));

        if (emails.length === 0) {
            return res.render('admin/email', { error: 'No valid email recipients found for this group.' });
        }

        // Setup Transporter
        const transporter = nodemailer.createTransport({
            service: 'gmail', // Standard support for Gmail; for others use host/port
            auth: {
                user: process.env.EMAIL_USER,
                pass: process.env.EMAIL_PASS
            }
        });

        const mailOptions = {
            from: process.env.EMAIL_USER || 'admin@waqarschool.edu.pk',
            bcc: emails, // Use BCC for bulk email privacy
            subject: subject,
            text: message
        };

        try {
            if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
                console.log('>>> EMAIL SIMULATION (Missing Credentials) <<<');
                console.log(`Recipients (${emails.length}):`, emails);
                console.log(`Subject: ${subject}`);
                console.log(`Body: ${message}`);
                return res.render('admin/email', { error: 'Email feature requires EMAIL_USER and EMAIL_PASS in .env file. (Check console for simulated output)' });
            }

            await transporter.sendMail(mailOptions);
            res.render('admin/email', { success: `Email sent successfully to ${emails.length} recipients.` });
        } catch (error) {
            console.error('Email Send Error:', error);
            res.render('admin/email', { error: 'Failed to send email. Error: ' + error.message });
        }
    });
});

// ============================================================
// STATIC ROUTES (CMS REMOVED)
// ============================================================

app.get('/', (req, res) => {
    res.render('index', { user: req.session.user, isHomepage: true });
});
// ============================================================

// Helper: Haversine Distance Calculation (km)
function calculateDistance(lat1, lon1, lat2, lon2) {
    if (!lat1 || !lon1 || !lat2 || !lon2) return 0;
    const R = 6371; // Radius of the earth in km
    const dLat = (lat2 - lat1) * (Math.PI / 180);
    const dLon = (lon2 - lon1) * (Math.PI / 180);
    const a =
        Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Math.round(R * c); // Distance in km
}

// SECURITY LOGGING FUNCTION
async function logSecurityEvent(req, user, action) {
    try {
        let clientIp = requestIp.getClientIp(req) || req.ip;

        // 🛡️ Determine Location: Priority GPS > GeoIP
        let latitude = null;
        let longitude = null;
        let country = 'Unknown';
        let city = 'Unknown';

        // 1. GeoIP Lookup (Default labels)
        const geo = geoip.lookup(clientIp);
        if (geo) {
            latitude = geo.ll[0];
            longitude = geo.ll[1];
            country = geo.country;
            city = geo.city;
        }

        // 2. Exact GPS Override (Captured from Browser)
        if (req.body && req.body.latitude && req.body.longitude) {
            latitude = parseFloat(req.body.latitude);
            longitude = parseFloat(req.body.longitude);
            if (req.body.city) city = req.body.city;
            if (req.body.country) country = req.body.country;
        }

        // 🧪 Handle Localhost fallback if no location detected
        if ((clientIp === '::1' || clientIp === '127.0.0.1' || clientIp.includes('unknown')) && !latitude) {
            // Only if GPS is also missing, we can keep a test fallback or set to null
            // We'll set a generic 'Local Access' label instead of 'Simulated'
            city = 'Local Access';
            country = 'Local Network';
        }

        const ua = UAParser(req.headers['user-agent']);
        const currentLat = latitude;
        const currentLon = longitude;
        let distance = 0;
        let risk = 'Low';

        // Check distance from last login if available
        if (user && user.last_latitude && currentLat) {
            distance = calculateDistance(user.last_latitude, user.last_longitude, currentLat, currentLon);
            if (distance > 500) risk = 'High'; // >500km is suspicious
            else if (distance > 50) risk = 'Medium';
        }

        // Detect Failed Login Spam
        if (action === 'FAILED_LOGIN') risk = 'Medium';

        // Prepare Insert
        const logData = {
            user_id: user ? user.id : null,
            campus_id: user ? user.campus_id : (req.session.campus ? req.session.campus.id : 1),
            role: user ? user.role : 'guest',
            action: action,
            ip_address: clientIp,
            latitude: currentLat,
            longitude: currentLon,
            country: country,
            city: city,
            device: ua.device.type || 'Desktop',
            browser: `${ua.browser.name} ${ua.browser.version}`,
            os: `${ua.os.name} ${ua.os.version}`,
            environment: (ua.device.type === 'mobile' || ua.device.type === 'tablet') ? 'Mobile' : 'Web',
            distance_from_last_login: distance,
            risk_level: risk
        };

        // 🔒 SECURITY FIX: Generate Integrity Hash (HMAC)
        const integrityData = `${logData.user_id}-${logData.action}-${logData.ip_address}-${logData.risk_level}`;
        logData.integrity_hash = crypto.createHmac('sha256', process.env.ENCRYPTION_KEY)
            .update(integrityData)
            .digest('hex');

        const sql = `INSERT INTO audit_logs SET ?`;
        db.query(sql, logData, (err) => {
            if (err) console.error("Audit Log Error:", err.message);

            // 🧹 AUTO-CLEANUP: Keep only the latest 2000 records
            const cleanupQuery = `
                DELETE FROM audit_logs 
                WHERE id NOT IN (
                    SELECT id FROM (
                        SELECT id FROM audit_logs ORDER BY id DESC LIMIT 2000
                    ) AS keep_rows
                )
            `;
            db.query(cleanupQuery, (cleanErr) => {
                if (cleanErr) console.error("Log Cleanup Error:", cleanErr.message);
            });
        });

        // Update User Last Known Location (Only on successful login actions)
        if (user && user.id && (action === 'LOGIN' || action === 'NEW_DEVICE_LOGIN')) {
            db.query("UPDATE users SET last_latitude = ?, last_longitude = ?, last_login_at = NOW() WHERE id = ?",
                [currentLat, currentLon, user.id]);
        }

    } catch (e) {
        console.error("Security Logging Exception:", e);
    }
}

// ==========================================
// 🎨 CMS & WEBSITE MANAGEMENT ROUTES
// ==========================================

// 1. CMS Settings
app.get('/admin/cms/settings', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    db.query('SELECT * FROM cms_settings', (err, results) => {
        const settings = {};
        if (results) results.forEach(r => settings[r.setting_key] = r.setting_value);
        res.render('admin/cms/settings', { settings });
    });
});

app.post('/admin/cms/settings/update', upload.fields([{ name: 'hero_image', maxCount: 1 }, { name: 'site_logo', maxCount: 1 }]), csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    const updates = { ...req.body };
    delete updates._csrf;

    if (req.files) {
        if (req.files['hero_image'] && req.files['hero_image'][0]) {
            updates.hero_image = '/uploads/' + req.files['hero_image'][0].filename;
        }
        if (req.files['site_logo'] && req.files['site_logo'][0]) {
            updates.site_logo = '/uploads/' + req.files['site_logo'][0].filename;
        }
    }

    const queries = Object.keys(updates).map(key => {
        return new Promise((resolve, reject) => {
            db.query('INSERT INTO cms_settings (setting_key, setting_value) VALUES (?, ?) ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value)',
                [key, updates[key]], (err) => err ? reject(err) : resolve());
        });
    });

    Promise.all(queries)
        .then(() => res.redirect('/admin/cms/settings?success=true'))
        .catch(err => res.status(500).send(err.message));
});

// 2. CMS Theme
app.get('/admin/cms/theme', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    db.query('SELECT * FROM cms_themes WHERE is_active = TRUE LIMIT 1', (err, results) => {
        let theme = results && results.length > 0 ? results[0] : {
            colors: '{"primary":"#1e3a8a","secondary":"#3b82f6","accent":"#60a5fa"}',
            fonts: '{"heading":"Inter, sans-serif","body":"Inter, sans-serif"}',
            button_styles: '{"borderRadius":"8px","shadow":"none"}'
        };
        try { if (typeof theme.colors === 'string') theme.colors = JSON.parse(theme.colors); } catch (e) { }
        try { if (typeof theme.fonts === 'string') theme.fonts = JSON.parse(theme.fonts); } catch (e) { }
        try { if (typeof theme.button_styles === 'string') theme.button_styles = JSON.parse(theme.button_styles); } catch (e) { }
        res.render('admin/cms/theme', { theme, csrfToken: res.locals.csrfToken });
    });
});

app.post('/admin/cms/theme/update', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const body = req.body;
    const colorsInput = body.colors || {};
    const fontsInput = body.fonts || {};
    const btnInput = body.button_styles || {};

    const colors = JSON.stringify({
        primary: colorsInput.primary || '#1e3a8a',
        secondary: colorsInput.secondary || '#3b82f6',
        accent: colorsInput.accent || '#60a5fa'
    });
    const fonts = JSON.stringify({
        heading: fontsInput.heading || 'Inter, sans-serif',
        body: fontsInput.body || 'Inter, sans-serif'
    });
    const buttonStyles = JSON.stringify({
        borderRadius: btnInput.borderRadius || '8px',
        shadow: btnInput.shadow || 'none'
    });

    db.query('UPDATE cms_themes SET colors = ?, fonts = ?, button_styles = ? WHERE is_active = TRUE', [colors, fonts, buttonStyles], (err) => {
        if (err) console.error(err);
        res.redirect('/admin/cms/theme?success=true');
    });
});

// 3. CMS Pages
app.get('/admin/cms/pages', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    db.query('SELECT * FROM cms_pages ORDER BY created_at DESC', (err, pages) => {
        res.render('admin/cms/pages', { pages: pages || [] });
    });
});

app.get('/admin/cms/pages/edit/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');

    if (req.params.id === 'new') {
        res.render('admin/cms/page_editor', { page: {}, sections: [], csrfToken: res.locals.csrfToken });
    } else {
        db.query('SELECT * FROM cms_pages WHERE id = ?', [req.params.id], (err, result) => {
            if (err || result.length === 0) return res.redirect('/admin/cms/pages');
            db.query('SELECT * FROM cms_sections WHERE page_id = ? ORDER BY display_order', [req.params.id], (err2, sections) => {
                res.render('admin/cms/page_editor', { page: result[0], sections: sections || [], csrfToken: res.locals.csrfToken });
            });
        });
    }
});

app.post('/admin/cms/pages/save', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { id, title, slug, content, meta_description } = req.body;

    if (id) {
        db.query('UPDATE cms_pages SET title=?, slug=?, content=?, meta_description=? WHERE id=?',
            [title, slug, content, meta_description, id], (err) => {
                if (err) console.error(err);
                res.redirect('/admin/cms/pages');
            });
    } else {
        db.query('INSERT INTO cms_pages (title, slug, content, meta_description, is_published) VALUES (?, ?, ?, ?, 1)',
            [title, slug, content, meta_description], (err) => {
                if (err) console.error(err);
                res.redirect('/admin/cms/pages');
            });
    }
});

app.get('/admin/cms/pages/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    db.query('DELETE FROM cms_pages WHERE id = ?', [req.params.id], (err) => {
        res.redirect('/admin/cms/pages');
    });
});

// 4. CMS Menus
app.get('/admin/cms/menus', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    db.query('SELECT * FROM cms_menu_items ORDER BY menu_location, display_order', (err, menus) => {
        res.render('admin/cms/menus', { menus: menus || [] });
    });
});

app.post('/admin/cms/menus/add', csrfProtection, (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    const { label, url, menu_location, display_order } = req.body;
    db.query('INSERT INTO cms_menu_items (label, url, menu_location, display_order, enabled) VALUES (?, ?, ?, ?, 1)',
        [label, url, menu_location, display_order], (err) => {
            if (err) console.error(err);
            res.redirect('/admin/cms/menus');
        });
});

app.get('/admin/cms/menus/delete/:id', (req, res) => {
    if (!req.session.user || req.session.user.role !== 'admin') return res.redirect('/login');
    db.query('DELETE FROM cms_menu_items WHERE id = ?', [req.params.id], (err) => {
        res.redirect('/admin/cms/menus');
    });
});

// 5. Public Dynamic Page Route (Frontend)
app.get('/p/:slug', (req, res) => {
    const slug = req.params.slug;
    db.query('SELECT * FROM cms_pages WHERE slug = ? AND is_published = 1', [slug], (err, results) => {
        if (err || results.length === 0) return res.status(404).render('404');
        res.locals.page = results[0]; // Set page title/content for header/meta
        res.render('dynamic_page', { pageContent: results[0] });
    });
});

// 404 Page Not Found (Catch-all)
app.use((req, res) => {
    res.status(404).render('404');
});

// 🔒 Global Error Handler (MUST BE LAST)
app.use((err, req, res, next) => {
    if (err.code === 'EBADCSRFTOKEN') {
        console.warn(`[SECURITY] CSRF Token Validation Failed for ${req.path}`);

        // Handle based on context
        if (req.path.startsWith('/auth/login') || req.path.startsWith('/login')) {
            return res.redirect('/login?error=session_expired');
        }

        return res.status(403).send(`
            <div style="font-family: sans-serif; text-align: center; padding: 50px; background: #f8fafc; min-height: 100vh; display: flex; flex-direction: column; align-items: center; justify-content: center;">
                <div style="background: white; padding: 40px; border-radius: 24px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1); max-width: 500px;">
                    <div style="width: 80px; height: 80px; background: #fee2e2; color: #ef4444; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 20px; font-size: 2rem;">
                        <i class="fas fa-shield-alt"></i>
                    </div>
                    <h1 style="color: #0f172a; margin-bottom: 10px; font-weight: 800;">Security Session Expired</h1>
                    <p style="color: #64748b; margin-bottom: 30px; line-height: 1.6;">For your protection, the request was blocked because the security token has expired. This happens if you leave a page open for too long or restart the server.</p>
                    <div style="display: flex; gap: 15px; justify-content: center;">
                        <button onclick="window.location.reload()" style="padding: 12px 25px; background: #3b82f6; color: white; border: none; border-radius: 12px; cursor: pointer; font-weight: 700;">Refresh Page</button>
                        <button onclick="window.history.back()" style="padding: 12px 25px; background: #f1f5f9; color: #475569; border: none; border-radius: 12px; cursor: pointer; font-weight: 700;">Go Back</button>
                    </div>
                </div>
            </div>
        `);
    }

    // Final Safe Catch
    if (res.headersSent) return next(err);
    console.error("[CRITICAL ERROR]", err);
    res.status(500).send("Something went wrong!");
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log(`> Update Loaded: Security & Audits Active.`);
});
