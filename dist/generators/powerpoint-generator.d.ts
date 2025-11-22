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
    defineMasterSlide(filename: string, masterSlide: {
        name: string;
        background?: {
            color?: string;
            image?: string;
        };
        placeholders?: Array<{
            type: 'title' | 'body' | 'footer' | 'slideNumber' | 'date';
            x: number;
            y: number;
            w: number;
            h: number;
        }>;
        fonts?: {
            title?: string;
            body?: string;
        };
        colors?: {
            accent1?: string;
            accent2?: string;
            accent3?: string;
        };
    }): Promise<Buffer>;
    addHyperlinks(filename: string, slideNumber: number, links: Array<{
        text: string;
        url?: string;
        slide?: number;
        tooltip?: string;
    }>): Promise<Buffer>;
    addSections(filename: string, sections: Array<{
        name: string;
        startSlide: number;
    }>): Promise<Buffer>;
    addMorphTransition(filename: string, fromSlide: number, toSlide: number, duration?: number): Promise<Buffer>;
    addActionButtons(filename: string, slideNumber: number, buttons: Array<{
        text: string;
        action: 'nextSlide' | 'previousSlide' | 'firstSlide' | 'lastSlide' | 'endShow' | 'customSlide';
        targetSlide?: number;
        x: number;
        y: number;
        w?: number;
        h?: number;
    }>): Promise<Buffer>;
}
//# sourceMappingURL=powerpoint-generator.d.ts.map