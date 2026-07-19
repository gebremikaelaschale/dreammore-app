require('dotenv').config(); // ይህ .env ፋይሉን ያነብልሃል

const express = require('express');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const cors = require('cors');
const multer = require('multer');
const mongoose = require('mongoose');
const upload = multer({ dest: 'uploads/' });

const app = express();
app.use(cors());
app.use(express.json());

const progressClients = new Set();
let progressState = { percent: 0, status: 'Idle' };

const emailLogSchema = new mongoose.Schema({
    studentName: { type: String, required: true, trim: true },
    email: { type: String, required: true, trim: true, lowercase: true },
    course: { type: String, required: true, trim: true, lowercase: true },
    sentAt: { type: Date, default: Date.now }
});

emailLogSchema.index({ email: 1, course: 1 }, { unique: true });

const EmailLog = mongoose.model('EmailLog', emailLogSchema);

mongoose.connect(process.env.MONGODB_URI)
    .then(() => console.log('Database Connected Successfully'))
    .catch((err) => console.error('MongoDB connection error:', err.message));

const broadcastProgress = (percent, status) => {
    progressState = { percent, status };
    progressClients.forEach((send) => send(progressState));
};

app.get('/progress', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders?.();

    const send = (payload) => {
        res.write(`data: ${JSON.stringify(payload)}\n\n`);
    };

    progressClients.add(send);
    send(progressState);

    req.on('close', () => {
        progressClients.delete(send);
    });
});

app.post('/send-emails', upload.single('file'), async (req, res) => {
    try {
        const { courseName, selectedRecipients } = req.body;
        console.log('Received request for course:', req.body.courseName);
        const workbook = xlsx.readFile(req.file.path);
        const data = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        const selectedRecipientSet = new Set((JSON.parse(selectedRecipients || '[]')).map((value) => String(value || '').trim().toLowerCase()));

        const filtered = data.filter((row) => {
            const matchesCourse = String(row.Course || '').trim().toLowerCase() === String(courseName || '').trim().toLowerCase();
            const email = String(row.Email || '').trim().toLowerCase();
            const isSelectedRecipient = selectedRecipientSet.size === 0 || selectedRecipientSet.has(email);
            return matchesCourse && isSelectedRecipient;
        });
        const validRecipients = filtered.filter((row) => {
            const email = String(row.Email || '').trim();
            return email.includes('@') && email.includes('.com');
        });
        const invalidRecipients = filtered.filter((row) => {
            const email = String(row.Email || '').trim();
            return !email.includes('@') || !email.includes('.com');
        });
        const total = validRecipients.length;

        broadcastProgress(0, 'Preparing recipients...');

        if (total === 0) {
            broadcastProgress(100, 'No valid recipients found.');
            return res.json({ success: true, message: 'No valid recipients found.', sentCount: 0, duplicateCount: 0, sentTo: [], duplicates: [], invalidRecipients });
        }

        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS }
        });

        const sentTo = [];
        const duplicates = [];

        for (let index = 0; index < validRecipients.length; index++) {
            const user = validRecipients[index];
            const normalizedEmail = String(user.Email || '').trim().toLowerCase();
            const normalizedCourse = String(courseName || '').trim().toLowerCase();
            const existingRecord = await EmailLog.findOne({ email: normalizedEmail, course: normalizedCourse });

            if (existingRecord) {
                duplicates.push({
                    studentName: user.Name || 'Unknown',
                    email: normalizedEmail,
                    course: courseName,
                    reason: 'duplicate'
                });
            } else {
                console.log('Attempting to send email to:', user.Email);
                try {
                    await transporter.sendMail({
                        from: `DreamMore <${process.env.EMAIL_USER}>`,
                        to: normalizedEmail,
                        subject: 'Course Registration',
                        html: `
                        <div style="background-color:#f4f4f4;padding:40px 20px;font-family:Arial,Helvetica,sans-serif;">
                          <div style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 6px 18px rgba(0,0,0,0.08);">
                            <div style="background:#eaf4ff;padding:28px 24px 20px;text-align:center;">
                              <img src="https://i.postimg.cc/cHdnkM74/photo-2024-10-15-01-22-00.jpg" alt="DreamMore logo" style="max-width:180px;height:auto;display:block;margin:0 auto;" />
                            </div>
                            <div style="padding:30px 32px 32px;color:#1f2937;line-height:1.7;">
                              <h2 style="margin:0 0 10px;font-size:26px;color:#0f172a;">Welcome to the Future of Your Career!</h2>
                              <p style="margin:0 0 12px;font-size:16px;">
                                Congratulations ${user.Name || 'Student'}!
                              </p>
                              <p style="margin:0 0 12px;font-size:16px;">
                                We are thrilled to officially welcome you to the <strong>${courseName}</strong> program at DreamMore.
                                By choosing this path, you have taken a significant step towards mastering the skills needed for today’s competitive market.
                              </p>
                              <h3 style="margin:18px 0 10px;font-size:18px;color:#0f172a;">What’s Next?</h3>
                              <ul style="margin:0 0 16px 20px;padding:0;font-size:15px;color:#374151;">
                                <li>Review your course syllabus.</li>
                                <li>Prepare your workstation for the upcoming sessions.</li>
                                <li>Join our student community on Slack/Telegram.</li>
                              </ul>
                              <p style="margin:0 0 8px;font-size:15px;">
                                We are committed to providing you with the right work at the right time. Our team is here to support you every step of the way.
                              </p>
                              <p style="margin:18px 0 0;font-size:14px;color:#4b5563;">
                                Warm regards,<br />
                                <strong>DreamMore Team</strong>
                              </p>
                            </div>
                            <div style="background:#f8fafc;padding:16px 24px;text-align:center;font-size:13px;color:#6b7280;border-top:1px solid #e5e7eb;">
                              <p style="margin:0;">DreamMore - Right work at right time</p>
                            </div>
                          </div>
                        </div>
                    `
                    });

                    await EmailLog.create({
                        studentName: user.Name || 'Unknown',
                        email: normalizedEmail,
                        course: normalizedCourse,
                        sentAt: new Date()
                    });

                    sentTo.push({
                        studentName: user.Name || 'Unknown',
                        email: normalizedEmail,
                        course: courseName
                    });
                } catch (error) {
                    console.error('Nodemailer Error:', error);
                }
            }

            const percent = Math.round(((index + 1) / total) * 100);
            broadcastProgress(percent, `Processing ${index + 1} of ${total}...`);
        }

        broadcastProgress(100, 'Completed');
        res.json({
            success: true,
            message: `${sentTo.length} emails sent and ${duplicates.length} duplicates skipped.`,
            sentCount: sentTo.length,
            duplicateCount: duplicates.length,
            sentTo,
            duplicates,
            invalidRecipients
        });
    } catch (err) {
        broadcastProgress(0, 'Failed');
        res.status(500).json({ error: err.message });
    }
});

app.get('/history', async (req, res) => {
    try {
        const history = await EmailLog.find().sort({ sentAt: -1 });
        res.json({ success: true, history });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.delete('/history/clear', async (req, res) => {
    try {
        const result = await EmailLog.deleteMany({});
        res.json({
            success: true,
            message: 'All history records have been deleted successfully.',
            deletedCount: result.deletedCount || 0
        });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.delete('/history/delete-selected', async (req, res) => {
    try {
        const ids = Array.isArray(req.body?.ids) ? req.body.ids : [];

        if (!ids.length) {
            return res.status(400).json({ success: false, message: 'Please provide an array of IDs to delete.' });
        }

        const result = await EmailLog.deleteMany({ _id: { $in: ids } });
        res.json({
            success: true,
            message: 'Selected history records have been deleted successfully.',
            deletedCount: result.deletedCount || 0
        });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.listen(5000, () => console.log("Server running on 5000"));