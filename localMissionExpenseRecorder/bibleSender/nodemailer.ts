import * as nodemailer from 'nodemailer'

export const emailUser = process.env.SMTP_USER;
export const emailTransporter = nodemailer.createTransport({
    service: "Outlook365",
    auth: {
      user: emailUser,
      pass: process.env.SMTP_PASS,
    },
});