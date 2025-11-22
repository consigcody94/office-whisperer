/**
 * Outlook Generator - Email, calendar, contacts, tasks, and rules
 * Note: Uses nodemailer for email, generates .ics/.vcf files for calendar/contacts
 */
import type { OutlookSendEmailArgs, OutlookCreateMeetingArgs, OutlookAddContactArgs, OutlookCreateTaskArgs, OutlookSetRuleArgs } from '../types.js';
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
}
//# sourceMappingURL=outlook-generator.d.ts.map