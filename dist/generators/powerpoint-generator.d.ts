/**
 * PowerPoint Generator - Create presentations using PptxGenJS
 */
import type { PowerPointPresentationOptions, PowerPointSlide, PPTTransition, PPTAnimation } from '../types.js';
export declare class PowerPointGenerator {
    createPresentation(options: PowerPointPresentationOptions): Promise<Buffer>;
    addTransition(filename: string, transition: PPTTransition, slideNumber?: number): Promise<Buffer>;
    addAnimation(filename: string, slideNumber: number, animation: PPTAnimation, objectId?: string): Promise<Buffer>;
    addNotes(filename: string, slideNumber: number, notes: string): Promise<Buffer>;
    duplicateSlide(filename: string, slideNumber: number, position?: number): Promise<Buffer>;
    reorderSlides(filename: string, slideOrder: number[]): Promise<Buffer>;
    exportPDF(filename: string): Promise<Buffer>;
    addMedia(filename: string, slideNumber: number, mediaPath: string, mediaType: 'video' | 'audio', position?: {
        x: number;
        y: number;
    }, size?: {
        width: number;
        height: number;
    }): Promise<Buffer>;
    addSlide(filename: string, slide: PowerPointSlide, position?: number): Promise<Buffer>;
    private applyTheme;
    private addContent;
}
//# sourceMappingURL=powerpoint-generator.d.ts.map