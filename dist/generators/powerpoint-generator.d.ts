/**
 * PowerPoint Generator - Create presentations using PptxGenJS
 */
import type { PowerPointPresentationOptions, PowerPointSlide, PPTTransition, PPTAnimation, PPTSmartArtArgs, PPTInsertIconsArgs, PPTInsert3DModelArgs, PPTZoomArgs, PPTRecordingArgs, PPTLiveWebArgs, PPTDesignerArgs, PPTCollaborationArgs, PPTPresenterCoachArgs, PPTSubtitlesArgs, PPTInkAnnotationsArgs, PPTGridGuidesArgs, PPTCustomShowArgs, PPTAnimationPaneArgs, PPTSlideMasterAdvancedArgs, PPTThemeArgs, PPTTemplateArgs } from '../types.js';
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
    /**
     * Add SmartArt graphics to a slide
     * Note: PptxGenJS doesn't natively support SmartArt, so this creates a placeholder
     * with shapes and text representing the SmartArt structure
     *
     * @param args - SmartArt configuration including type, layout, items, and styling
     * @returns Buffer containing the modified presentation
     */
    addSmartArt(args: PPTSmartArtArgs): Promise<Buffer>;
    /**
     * Insert SVG icons from Microsoft's icon library
     * Note: Creates placeholder shapes representing icons with position and color
     *
     * @param args - Icon configuration including name, category, position, size, and color
     * @returns Buffer containing the modified presentation
     */
    insertIcons(args: PPTInsertIconsArgs): Promise<Buffer>;
    /**
     * Insert 3D models into presentation
     * Note: Creates placeholder with metadata since PptxGenJS doesn't support 3D models natively
     *
     * @param args - 3D model configuration including path, position, rotation, and animation
     * @returns Buffer containing the modified presentation
     */
    insert3DModels(args: PPTInsert3DModelArgs): Promise<Buffer>;
    /**
     * Add Zoom links (Summary Zoom, Slide Zoom, Section Zoom)
     * Note: Creates placeholder with visual preview representation
     *
     * @param args - Zoom configuration including type, target slides/sections, and display options
     * @returns Buffer containing the modified presentation
     */
    addZoom(args: PPTZoomArgs): Promise<Buffer>;
    /**
     * Configure screen recording and audio narration settings
     * Note: Stores metadata for recording configuration
     *
     * @param args - Recording settings including type, quality, slides, and capture options
     * @returns Buffer containing the presentation with recording metadata
     */
    configureRecording(args: PPTRecordingArgs): Promise<Buffer>;
    /**
     * Embed live web page in a slide
     * Note: Creates placeholder with URL metadata
     *
     * @param args - Web page embedding configuration including URL, position, size, and interaction settings
     * @returns Buffer containing the modified presentation
     */
    embedLiveWeb(args: PPTLiveWebArgs): Promise<Buffer>;
    /**
     * Apply PowerPoint Designer suggestions
     * Note: Stores designer preferences as metadata; actual AI suggestions require PowerPoint
     *
     * @param args - Designer preferences including style, layout, and color palette
     * @returns Buffer containing the presentation with designer metadata
     */
    applyDesigner(args: PPTDesignerArgs): Promise<Buffer>;
    /**
     * Add collaboration comments and @mentions
     * Note: Creates comment metadata; actual threaded comments require PowerPoint
     *
     * @param args - Comment configuration including text, author, position, mentions, and replies
     * @returns Buffer containing the presentation with comment metadata
     */
    addCollaborationComments(args: PPTCollaborationArgs): Promise<Buffer>;
    /**
     * Configure Presenter Coach rehearsal settings
     * Note: Stores Presenter Coach preferences as metadata
     *
     * @param args - Presenter Coach settings for feedback, pacing, filler words, etc.
     * @returns Buffer containing the presentation with Presenter Coach metadata
     */
    configurePresenterCoach(args: PPTPresenterCoachArgs): Promise<Buffer>;
    /**
     * Configure live subtitles and captions
     * Note: Stores subtitle/caption configuration as metadata
     *
     * @param args - Subtitle settings including language, position, styling, and translation options
     * @returns Buffer containing the presentation with subtitle metadata
     */
    configureSubtitles(args: PPTSubtitlesArgs): Promise<Buffer>;
    /**
     * Add digital ink drawings and annotations
     * Note: Creates visual representation of ink strokes
     *
     * @param args - Ink annotation configuration including pen/highlighter strokes with points and pressure
     * @returns Buffer containing the modified presentation
     */
    addInkAnnotations(args: PPTInkAnnotationsArgs): Promise<Buffer>;
    /**
     * Configure alignment grids and guides
     * Note: Stores grid and guide configuration as metadata
     *
     * @param args - Grid and guide settings including spacing, snap options, and guide positions
     * @returns Buffer containing the presentation with grid/guide metadata
     */
    configureGridGuides(args: PPTGridGuidesArgs): Promise<Buffer>;
    /**
     * Create custom slide shows for different audiences
     * Note: Stores custom show definitions as metadata
     *
     * @param args - Custom show configurations with names, slide selections, and descriptions
     * @returns Buffer containing the presentation with custom show metadata
     */
    createCustomShow(args: PPTCustomShowArgs): Promise<Buffer>;
    /**
     * Manage animation timing and order via Animation Pane
     * Note: Stores animation sequence metadata
     *
     * @param args - Animation pane configuration with effects, triggers, timing, and sequencing
     * @returns Buffer containing the modified presentation
     */
    manageAnimationPane(args: PPTAnimationPaneArgs): Promise<Buffer>;
    /**
     * Customize slide master with advanced settings
     * Note: Creates comprehensive slide master with fonts, colors, effects, and placeholders
     *
     * @param args - Advanced slide master configuration including theme, placeholders, and background
     * @returns Buffer containing the presentation with customized master
     */
    customizeSlideMaster(args: PPTSlideMasterAdvancedArgs): Promise<Buffer>;
    /**
     * Apply or customize presentation theme
     * Note: Applies theme settings including colors, fonts, and variants
     *
     * @param args - Theme configuration with name, variant, and custom colors/fonts
     * @returns Buffer containing the presentation with applied theme
     */
    applyPresentationTheme(args: PPTThemeArgs): Promise<Buffer>;
    /**
     * Save presentation as a template
     * Note: Creates template with placeholders and metadata for protected elements
     *
     * @param args - Template configuration with title, category, placeholders, and protected elements
     * @returns Buffer containing the template (.potx format)
     */
    saveAsTemplate(args: PPTTemplateArgs): Promise<Buffer>;
    private getSmartArtColors;
}
//# sourceMappingURL=powerpoint-generator.d.ts.map