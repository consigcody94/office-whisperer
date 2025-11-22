/**
 * PowerPoint Generator - Create presentations using PptxGenJS
 */
import type { PowerPointPresentationOptions, PowerPointSlide } from '../types.js';
export declare class PowerPointGenerator {
    createPresentation(options: PowerPointPresentationOptions): Promise<Buffer>;
    private applyTheme;
    private addContent;
    addSlide(filename: string, slide: PowerPointSlide, position?: number): Promise<Buffer>;
}
//# sourceMappingURL=powerpoint-generator.d.ts.map