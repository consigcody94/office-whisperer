/**
 * Outlook Generator - Email, calendar, contacts, tasks, and rules
 * Note: Uses nodemailer for email, generates .ics/.vcf files for calendar/contacts
 */
import type { OutlookSendEmailArgs, OutlookCreateMeetingArgs, OutlookAddContactArgs, OutlookCreateTaskArgs, OutlookSetRuleArgs, OutlookAttendee } from '../types.js';
export declare class OutlookGenerator {
    /**
     * Send email using SMTP (requires SMTP configuration)
     */
    sendEmail(args: OutlookSendEmailArgs): Promise<string>;
    /**
     * Create calendar meeting (generates .ics file)
     */
    createMeeting(args: OutlookCreateMeetingArgs): Promise<string>;
    /**
     * Add contact (generates .vcf vCard file)
     */
    addContact(args: OutlookAddContactArgs): Promise<string>;
    /**
     * Create Outlook task (returns structured JSON)
     */
    createTask(args: OutlookCreateTaskArgs): Promise<string>;
    /**
     * Create inbox rule (returns structured JSON)
     */
    setRule(args: OutlookSetRuleArgs): Promise<string>;
    private generateICS;
    private generateVCF;
    /**
     * Read emails from IMAP server
     * Note: Requires 'imap' package for actual implementation
     */
    readEmails(args: {
        folder?: string;
        limit?: number;
        unreadOnly?: boolean;
        since?: string;
        imapConfig?: {
            host: string;
            port: number;
            user: string;
            password: string;
            tls?: boolean;
        };
    }): Promise<string>;
    /**
     * Search emails by query
     */
    searchEmails(args: {
        query: string;
        searchIn?: ('subject' | 'from' | 'body' | 'to')[];
        folder?: string;
        limit?: number;
        since?: string;
        imapConfig?: {
            host: string;
            port: number;
            user: string;
            password: string;
            tls?: boolean;
        };
    }): Promise<string>;
    /**
     * Create recurring meeting (generates .ics file with RRULE)
     */
    createRecurringMeeting(args: {
        subject: string;
        startTime: string;
        endTime: string;
        recurrence: {
            frequency: 'daily' | 'weekly' | 'monthly' | 'yearly';
            interval?: number;
            daysOfWeek?: ('MO' | 'TU' | 'WE' | 'TH' | 'FR' | 'SA' | 'SU')[];
            until?: string;
            count?: number;
        };
        location?: string;
        attendees?: OutlookAttendee[];
        description?: string;
    }): Promise<string>;
    /**
     * Save email template
     */
    saveEmailTemplate(args: {
        name: string;
        subject: string;
        body: string;
        html?: boolean;
        placeholders?: string[];
    }): Promise<string>;
    /**
     * Mark emails as read/unread
     */
    markAsRead(args: {
        messageIds: string[];
        markAsRead: boolean;
        imapConfig?: {
            host: string;
            port: number;
            user: string;
            password: string;
            tls?: boolean;
        };
    }): Promise<string>;
    /**
     * Archive emails
     */
    archiveEmail(args: {
        messageIds: string[];
        archiveFolder?: string;
        imapConfig?: {
            host: string;
            port: number;
            user: string;
            password: string;
            tls?: boolean;
        };
    }): Promise<string>;
    /**
     * Get calendar view for date range
     */
    getCalendarView(args: {
        startDate: string;
        endDate: string;
        viewType: 'day' | 'week' | 'month' | 'agenda';
        outputFormat?: 'ics' | 'json';
    }): Promise<string>;
    /**
     * Search contacts
     */
    searchContacts(args: {
        query: string;
        searchIn?: ('name' | 'email' | 'company' | 'phone')[];
        outputFormat?: 'vcf' | 'json';
    }): Promise<string>;
}
//# sourceMappingURL=outlook-generator.d.ts.map