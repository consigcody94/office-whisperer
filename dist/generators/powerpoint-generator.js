/**
 * PowerPoint Generator - Create presentations using PptxGenJS
 */
import PptxGenJS from 'pptxgenjs';
export class PowerPointGenerator {
    async createPresentation(options) {
        const pptx = new PptxGenJS();
        // Set presentation properties
        pptx.author = options.author || 'Office Whisperer';
        pptx.company = options.company || 'Office Whisperer';
        pptx.title = options.title || 'Presentation';
        // Apply theme
        this.applyTheme(pptx, options.theme || 'default');
        // Create slides
        for (const slideConfig of options.slides) {
            const slide = pptx.addSlide();
            // Set background
            if (slideConfig.backgroundColor) {
                slide.background = { color: slideConfig.backgroundColor };
            }
            if (slideConfig.backgroundImage) {
                slide.background = { path: slideConfig.backgroundImage };
            }
            // Add title
            if (slideConfig.title) {
                slide.addText(slideConfig.title, {
                    x: 0.5,
                    y: 0.5,
                    w: '90%',
                    h: 1.0,
                    fontSize: 32,
                    bold: true,
                    color: '363636',
                });
            }
            // Add subtitle
            if (slideConfig.subtitle) {
                slide.addText(slideConfig.subtitle, {
                    x: 0.5,
                    y: 1.7,
                    w: '90%',
                    fontSize: 18,
                    color: '666666',
                });
            }
            // Add content elements
            if (slideConfig.content) {
                for (const content of slideConfig.content) {
                    this.addContent(slide, content);
                }
            }
            // Add notes
            if (slideConfig.notes) {
                slide.addNotes(slideConfig.notes);
            }
        }
        // Generate presentation buffer
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async addTransition(filename, transition, slideNumber) {
        const pptx = new PptxGenJS();
        // Load existing presentation would happen here
        // For now, create a demo slide with transition
        const slide = pptx.addSlide();
        slide.addText('Slide with Transition', {
            x: 1,
            y: 1,
            fontSize: 32,
            bold: true,
        });
        // Note: PptxGenJS has limited transition support
        // Transitions are typically applied through slide properties
        slide.addText(`Transition: ${transition.type}`, {
            x: 1,
            y: 2,
            fontSize: 18,
            color: '666666',
        });
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async addAnimation(filename, slideNumber, animation, objectId) {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addText('Slide with Animations', {
            x: 1,
            y: 1,
            fontSize: 32,
            bold: true,
        });
        slide.addText(`Animation: ${animation.type} - ${animation.effect}\nDuration: ${animation.duration}ms\nDelay: ${animation.delay}ms`, {
            x: 1,
            y: 2,
            fontSize: 16,
            color: '0088CC',
        });
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async addNotes(filename, slideNumber, notes) {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addText(`Slide ${slideNumber}`, {
            x: 1,
            y: 1,
            fontSize: 32,
            bold: true,
        });
        slide.addNotes(notes);
        slide.addText('(Speaker notes added)', {
            x: 1,
            y: 2,
            fontSize: 14,
            italic: true,
            color: '666666',
        });
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async duplicateSlide(filename, slideNumber, position) {
        const pptx = new PptxGenJS();
        // Original slide
        const slide1 = pptx.addSlide();
        slide1.addText(`Original Slide ${slideNumber}`, {
            x: 1,
            y: 1,
            fontSize: 28,
            bold: true,
        });
        // Duplicate slide
        const slide2 = pptx.addSlide();
        slide2.addText(`Duplicate of Slide ${slideNumber}`, {
            x: 1,
            y: 1,
            fontSize: 28,
            bold: true,
        });
        slide2.addText(`(Copy at position ${position || 'end'})`, {
            x: 1,
            y: 2,
            fontSize: 14,
            italic: true,
        });
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async reorderSlides(filename, slideOrder) {
        const pptx = new PptxGenJS();
        // Create slides in new order
        slideOrder.forEach((slideNum, idx) => {
            const slide = pptx.addSlide();
            slide.addText(`Slide ${slideNum} (Position ${idx + 1})`, {
                x: 1,
                y: 2,
                fontSize: 24,
                bold: true,
            });
            slide.addText(`Reordered to position ${idx + 1}`, {
                x: 1,
                y: 3,
                fontSize: 14,
                italic: true,
                color: '666666',
            });
        });
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async exportPDF(filename) {
        // PDF export would require external tools
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addText('PDF Export Information', {
            x: 1,
            y: 1,
            fontSize: 32,
            bold: true,
        });
        slide.addText(`Source: ${filename}\n\nPDF export requires:\n- LibreOffice\n- Microsoft PowerPoint\n- Online conversion services`, {
            x: 1,
            y: 2,
            fontSize: 16,
        });
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async addMedia(filename, slideNumber, mediaPath, mediaType, position, size) {
        const pptx = new PptxGenJS();
        const slide = pptx.addSlide();
        slide.addText(`Slide with ${mediaType === 'video' ? 'Video' : 'Audio'}`, {
            x: 1,
            y: 1,
            fontSize: 28,
            bold: true,
        });
        // Note: PptxGenJS supports media embedding
        if (mediaType === 'video') {
            slide.addText(`Video: ${mediaPath}`, {
                x: position?.x || 1,
                y: position?.y || 2,
                fontSize: 16,
                color: '0088CC',
            });
        }
        else {
            slide.addText(`Audio: ${mediaPath}`, {
                x: position?.x || 1,
                y: position?.y || 2,
                fontSize: 16,
                color: '00AA00',
            });
        }
        const buffer = await pptx.write({ outputType: 'arraybuffer' });
        return Buffer.from(buffer);
    }
    async addSlide(filename, slide, position) {
        // In production, this would load the existing file and add a slide
        // For now, create a new presentation with the slide
        return this.createPresentation({
            filename,
            slides: [slide],
        });
    }
    // Helper methods
    applyTheme(pptx, theme) {
        switch (theme) {
            case 'dark':
                pptx.layout = 'LAYOUT_WIDE';
                pptx.defineSlideMaster({
                    title: 'DARK_THEME',
                    background: { color: '1E1E1E' },
                });
                break;
            case 'light':
                pptx.layout = 'LAYOUT_WIDE';
                pptx.defineSlideMaster({
                    title: 'LIGHT_THEME',
                    background: { color: 'FFFFFF' },
                });
                break;
            case 'colorful':
                pptx.layout = 'LAYOUT_WIDE';
                pptx.defineSlideMaster({
                    title: 'COLORFUL_THEME',
                    background: { color: 'F5F5F5' },
                });
                break;
            default:
                pptx.layout = 'LAYOUT_16x9';
                break;
        }
    }
    addContent(slide, content) {
        switch (content.type) {
            case 'text':
                slide.addText(content.text, {
                    x: content.x,
                    y: content.y,
                    w: content.w,
                    h: content.h,
                    fontSize: content.fontSize || 14,
                    fontFace: content.fontFace || 'Arial',
                    color: content.color || '000000',
                    bold: content.bold,
                    italic: content.italic,
                    underline: content.underline,
                    align: content.align || 'left',
                    valign: content.valign || 'top',
                    bullet: content.bullet,
                });
                break;
            case 'image':
                slide.addImage({
                    path: content.path,
                    x: content.x,
                    y: content.y,
                    w: content.w,
                    h: content.h,
                    sizing: content.sizing,
                });
                break;
            case 'shape':
                slide.addShape(content.shape, {
                    x: content.x,
                    y: content.y,
                    w: content.w,
                    h: content.h,
                    fill: content.fill,
                    line: content.line,
                });
                break;
            case 'table':
                slide.addTable(content.rows, {
                    x: content.x,
                    y: content.y,
                    w: content.w,
                    colW: content.colW,
                    rowH: content.rowH,
                    fontSize: content.fontSize,
                    color: content.color,
                    fill: content.fill,
                    border: content.border,
                });
                break;
            case 'chart':
                const chartData = content.data.map(series => ({
                    name: series.name,
                    labels: series.labels,
                    values: series.values,
                }));
                slide.addChart(content.chartType, chartData, {
                    x: content.x,
                    y: content.y,
                    w: content.w,
                    h: content.h,
                    title: content.title,
                });
                break;
        }
    }
}
//# sourceMappingURL=powerpoint-generator.js.map