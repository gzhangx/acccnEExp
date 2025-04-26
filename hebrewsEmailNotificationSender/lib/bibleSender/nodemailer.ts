import * as nodemailer from 'nodemailer'

export const emailUser = process.env.SMTP_USER || 'SMTP_USER Not Set';
export const emailTransporter = nodemailer.createTransport({
    service: "Outlook365",
    auth: {
      user: emailUser,
      pass: process.env.SMTP_PASS,
    },
});