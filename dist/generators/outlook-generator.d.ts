/**
 * Outlook Generator - Email, calendar, contacts, tasks, and rules
 * Note: Uses nodemailer for email, generates .ics/.vcf files for calendar/contacts
 */
import type { OutlookSendEmailArgs, OutlookCreateMeetingArgs, OutlookAddContactArgs, OutlookCreateTaskArgs, OutlookSetRuleArgs, OutlookAttendee, OutlookReadFullEmailArgs, OutlookDeleteEmailArgs, OutlookMoveEmailArgs, OutlookCreateFolderArgs, OutlookSharedMailboxArgs, OutlookDelegateAccessArgs, OutlookOutOfOfficeArgs, OutlookNotesArgs, OutlookJournalArgs, OutlookRSSFeedArgs, OutlookDataFileArgs, OutlookQuickStepsArgs, OutlookConversationViewArgs, OutlookCleanupArgs, OutlookIgnoreConversationArgs, OutlookFlagEmailArgs, OutlookCategoryArgs, OutlookSignatureArgs, OutlookAutoCompleteArgs, OutlookMailMergeAdvancedArgs } from '../types.js';
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
    /**
     * Read full email with attachments, headers, and raw MIME content via IMAP
     */
    readFullEmail(args: OutlookReadFullEmailArgs): Promise<string>;
    /**
     * Delete emails via IMAP (permanent or mark for deletion)
     */
    deleteEmail(args: OutlookDeleteEmailArgs): Promise<string>;
    /**
     * Move emails between IMAP folders
     */
    moveEmail(args: OutlookMoveEmailArgs): Promise<string>;
    /**
     * Create nested IMAP folder structure
     */
    createFolder(args: OutlookCreateFolderArgs): Promise<string>;
    /**
     * Access shared mailboxes and delegate accounts
     */
    accessSharedMailbox(args: OutlookSharedMailboxArgs): Promise<string>;
    /**
     * Grant delegate permissions to other users (generates metadata file)
     */
    grantDelegateAccess(args: OutlookDelegateAccessArgs): Promise<string>;
    /**
     * Set automatic replies / out of office / vacation responder
     */
    setOutOfOffice(args: OutlookOutOfOfficeArgs): Promise<string>;
    /**
     * Create Outlook notes with color coding
     */
    createNotes(args: OutlookNotesArgs): Promise<string>;
    /**
     * Create journal entries for activity tracking
     */
    createJournalEntry(args: OutlookJournalArgs): Promise<string>;
    /**
     * Subscribe to and manage RSS feeds
     */
    manageRSSFeed(args: OutlookRSSFeedArgs): Promise<string>;
    /**
     * Manage PST/OST data files (metadata operations only)
     */
    manageDataFile(args: OutlookDataFileArgs): Promise<string>;
    /**
     * Create Quick Steps for email automation
     */
    createQuickStep(args: OutlookQuickStepsArgs): Promise<string>;
    /**
     * Configure conversation view settings
     */
    configureConversationView(args: OutlookConversationViewArgs): Promise<string>;
    /**
     * Clean up redundant messages in conversations
     */
    cleanupMessages(args: OutlookCleanupArgs): Promise<string>;
    /**
     * Ignore conversation threads
     */
    ignoreConversation(args: OutlookIgnoreConversationArgs): Promise<string>;
    /**
     * Flag emails with colors and due dates
     */
    flagEmail(args: OutlookFlagEmailArgs): Promise<string>;
    /**
     * Create and apply color categories
     */
    manageCategories(args: OutlookCategoryArgs): Promise<string>;
    /**
     * Create HTML email signatures with images and formatting
     */
    createSignature(args: OutlookSignatureArgs): Promise<string>;
    /**
     * Manage autocomplete nickname cache
     */
    manageAutoComplete(args: OutlookAutoCompleteArgs): Promise<string>;
    /**
     * Advanced mail merge with filters and conditional content
     */
    advancedMailMerge(args: OutlookMailMergeAdvancedArgs): Promise<string>;
    /**
     * Generate OPML file for RSS feeds
     */
    private generateOPML;
    /**
     * Evaluate simple conditions for mail merge
     */
    private evaluateCondition;
}
//# sourceMappingURL=outlook-generator.d.ts.map