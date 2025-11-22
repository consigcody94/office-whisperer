/**
 * Outlook Generator - Email, calendar, contacts, tasks, and rules
 * Note: Uses nodemailer for email, generates .ics/.vcf files for calendar/contacts
 */

import nodemailer from 'nodemailer';
import type {
  OutlookSendEmailArgs,
  OutlookCreateMeetingArgs,
  OutlookAddContactArgs,
  OutlookCreateTaskArgs,
  OutlookSetRuleArgs,
  OutlookAttendee,
} from '../types.js';

export class OutlookGenerator {
  /**
   * Send email using SMTP (requires SMTP configuration)
   */
  async sendEmail(args: OutlookSendEmailArgs): Promise<string> {
    if (!args.smtpConfig) {
      return 'Email configuration:\n' +
        `To: ${Array.isArray(args.to) ? args.to.join(', ') : args.to}\n` +
        `Subject: ${args.subject}\n` +
        `Body: ${args.body.substring(0, 100)}...\n\n` +
        'Note: SMTP configuration required for actual sending.\n' +
        'Provide smtpConfig with host, port, and auth credentials.';
    }

    const transporter = nodemailer.createTransport({
      host: args.smtpConfig.host,
      port: args.smtpConfig.port,
      secure: args.smtpConfig.secure !== false,
      auth: args.smtpConfig.auth,
    });

    const mailOptions: any = {
      from: args.smtpConfig.auth?.user || 'noreply@officewhisperer.com',
      to: Array.isArray(args.to) ? args.to.join(', ') : args.to,
      subject: args.subject,
      [args.html ? 'html' : 'text']: args.body,
    };

    if (args.cc) {
      mailOptions.cc = Array.isArray(args.cc) ? args.cc.join(', ') : args.cc;
    }

    if (args.bcc) {
      mailOptions.bcc = Array.isArray(args.bcc) ? args.bcc.join(', ') : args.bcc;
    }

    if (args.attachments) {
      mailOptions.attachments = args.attachments.map(att => ({
        filename: att.filename,
        path: att.path,
        content: att.content,
      }));
    }

    if (args.priority) {
      mailOptions.priority = args.priority;
    }

    try {
      const info = await transporter.sendMail(mailOptions);
      return `Email sent successfully!\n` +
        `Message ID: ${info.messageId}\n` +
        `To: ${mailOptions.to}\n` +
        `Subject: ${args.subject}\n` +
        `${args.attachments ? `Attachments: ${args.attachments.length}\n` : ''}` +
        `Status: Delivered`;
    } catch (error) {
      throw new Error(`Failed to send email: ${error}`);
    }
  }

  /**
   * Create calendar meeting (generates .ics file)
   */
  async createMeeting(args: OutlookCreateMeetingArgs): Promise<string> {
    const startDate = new Date(args.startTime);
    const endDate = new Date(args.endTime);

    // Generate ICS (iCalendar) format
    const ics = this.generateICS({
      subject: args.subject,
      startDate,
      endDate,
      location: args.location,
      description: args.description,
      attendees: args.attendees,
      reminder: args.reminder,
    });

    return ics;
  }

  /**
   * Add contact (generates .vcf vCard file)
   */
  async addContact(args: OutlookAddContactArgs): Promise<string> {
    // Generate VCF (vCard) format
    const vcf = this.generateVCF({
      firstName: args.firstName,
      lastName: args.lastName,
      email: args.email,
      phone: args.phone,
      company: args.company,
      jobTitle: args.jobTitle,
      address: args.address,
    });

    return vcf;
  }

  /**
   * Create Outlook task (returns structured JSON)
   */
  async createTask(args: OutlookCreateTaskArgs): Promise<string> {
    const task = {
      subject: args.subject,
      dueDate: args.dueDate,
      priority: args.priority || 'normal',
      status: args.status || 'notStarted',
      category: args.category,
      reminder: args.reminder,
      notes: args.notes,
      createdAt: new Date().toISOString(),
    };

    return JSON.stringify(task, null, 2);
  }

  /**
   * Create inbox rule (returns structured JSON)
   */
  async setRule(args: OutlookSetRuleArgs): Promise<string> {
    const rule = {
      name: args.name,
      enabled: true,
      conditions: args.conditions.map(c => ({
        type: c.type,
        operator: c.operator || 'contains',
        value: c.value,
      })),
      actions: args.actions.map(a => ({
        type: a.type,
        value: a.value,
      })),
      createdAt: new Date().toISOString(),
    };

    return JSON.stringify(rule, null, 2);
  }

  // Helper methods
  private generateICS(params: {
    subject: string;
    startDate: Date;
    endDate: Date;
    location?: string;
    description?: string;
    attendees?: OutlookAttendee[];
    reminder?: number;
  }): string {
    const formatDate = (date: Date) => {
      return date.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
    };

    const uid = `${Date.now()}@officewhisperer.com`;
    const now = formatDate(new Date());
    const start = formatDate(params.startDate);
    const end = formatDate(params.endDate);

    let ics = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Office Whisperer//EN',
      'CALSCALE:GREGORIAN',
      'METHOD:REQUEST',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `DTSTART:${start}`,
      `DTEND:${end}`,
      `SUMMARY:${params.subject}`,
    ];

    if (params.location) {
      ics.push(`LOCATION:${params.location}`);
    }

    if (params.description) {
      ics.push(`DESCRIPTION:${params.description.replace(/\n/g, '\\n')}`);
    }

    if (params.attendees) {
      params.attendees.forEach(attendee => {
        const role = attendee.required !== false ? 'REQ-PARTICIPANT' : 'OPT-PARTICIPANT';
        const name = attendee.name || attendee.email;
        ics.push(`ATTENDEE;ROLE=${role};CN=${name}:mailto:${attendee.email}`);
      });
    }

    if (params.reminder) {
      ics.push('BEGIN:VALARM');
      ics.push('ACTION:DISPLAY');
      ics.push(`DESCRIPTION:Reminder`);
      ics.push(`TRIGGER:-PT${params.reminder}M`);
      ics.push('END:VALARM');
    }

    ics.push('STATUS:CONFIRMED');
    ics.push('SEQUENCE:0');
    ics.push('END:VEVENT');
    ics.push('END:VCALENDAR');

    return ics.join('\r\n');
  }

  private generateVCF(params: {
    firstName: string;
    lastName: string;
    email?: string;
    phone?: string;
    company?: string;
    jobTitle?: string;
    address?: string;
  }): string {
    let vcf = [
      'BEGIN:VCARD',
      'VERSION:3.0',
      `N:${params.lastName};${params.firstName};;;`,
      `FN:${params.firstName} ${params.lastName}`,
    ];

    if (params.email) {
      vcf.push(`EMAIL;TYPE=INTERNET:${params.email}`);
    }

    if (params.phone) {
      vcf.push(`TEL;TYPE=WORK,VOICE:${params.phone}`);
    }

    if (params.company) {
      vcf.push(`ORG:${params.company}`);
    }

    if (params.jobTitle) {
      vcf.push(`TITLE:${params.jobTitle}`);
    }

    if (params.address) {
      vcf.push(`ADR;TYPE=WORK:;;${params.address};;;;`);
    }

    vcf.push(`REV:${new Date().toISOString()}`);
    vcf.push('END:VCARD');

    return vcf.join('\r\n');
  }
}
