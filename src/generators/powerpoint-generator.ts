/**
 * PowerPoint Generator - Create presentations using PptxGenJS
 */

import PptxGenJS from 'pptxgenjs';
import type {
  PowerPointPresentationOptions,
  PowerPointSlide,
  PowerPointContent,
  PPTTransition,
  PPTAnimation,
} from '../types.js';

export class PowerPointGenerator {
  async createPresentation(options: PowerPointPresentationOptions): Promise<Buffer> {
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
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addTransition(
    filename: string,
    transition: PPTTransition,
    slideNumber?: number
  ): Promise<Buffer> {
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
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addAnimation(
    filename: string,
    slideNumber: number,
    animation: PPTAnimation,
    objectId?: string
  ): Promise<Buffer> {
    const pptx = new PptxGenJS();

    const slide = pptx.addSlide();
    slide.addText('Slide with Animations', {
      x: 1,
      y: 1,
      fontSize: 32,
      bold: true,
    });

    slide.addText(
      `Animation: ${animation.type} - ${animation.effect}\nDuration: ${animation.duration}ms\nDelay: ${animation.delay}ms`,
      {
        x: 1,
        y: 2,
        fontSize: 16,
        color: '0088CC',
      }
    );

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addNotes(
    filename: string,
    slideNumber: number,
    notes: string
  ): Promise<Buffer> {
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
    return Buffer.from(buffer as ArrayBuffer);
  }

  async duplicateSlide(
    filename: string,
    slideNumber: number,
    position?: number
  ): Promise<Buffer> {
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
    return Buffer.from(buffer as ArrayBuffer);
  }

  async reorderSlides(
    filename: string,
    slideOrder: number[]
  ): Promise<Buffer> {
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
    return Buffer.from(buffer as ArrayBuffer);
  }

  async exportPDF(filename: string): Promise<Buffer> {
    // PDF export would require external tools
    const pptx = new PptxGenJS();

    const slide = pptx.addSlide();
    slide.addText('PDF Export Information', {
      x: 1,
      y: 1,
      fontSize: 32,
      bold: true,
    });

    slide.addText(
      `Source: ${filename}\n\nPDF export requires:\n- LibreOffice\n- Microsoft PowerPoint\n- Online conversion services`,
      {
        x: 1,
        y: 2,
        fontSize: 16,
      }
    );

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addMedia(
    filename: string,
    slideNumber: number,
    mediaPath: string,
    mediaType: 'video' | 'audio',
    position?: { x: number; y: number },
    size?: { width: number; height: number }
  ): Promise<Buffer> {
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
    } else {
      slide.addText(`Audio: ${mediaPath}`, {
        x: position?.x || 1,
        y: position?.y || 2,
        fontSize: 16,
        color: '00AA00',
      });
    }

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addSlide(
    filename: string,
    slide: PowerPointSlide,
    position?: number
  ): Promise<Buffer> {
    // In production, this would load the existing file and add a slide
    // For now, create a new presentation with the slide
    return this.createPresentation({
      filename,
      slides: [slide],
    });
  }

  // Helper methods
  private applyTheme(pptx: PptxGenJS, theme: string): void {
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

  private addContent(slide: any, content: PowerPointContent): void {
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

  // ============================================================================
  // v3.0 Phase 1 Methods
  // ============================================================================

  async defineMasterSlide(
    filename: string,
    masterSlide: {
      name: string;
      background?: { color?: string; image?: string };
      placeholders?: Array<{
        type: 'title' | 'body' | 'footer' | 'slideNumber' | 'date';
        x: number;
        y: number;
        w: number;
        h: number;
      }>;
      fonts?: { title?: string; body?: string };
      colors?: { accent1?: string; accent2?: string; accent3?: string };
    }
  ): Promise<Buffer> {
    const pptx = new PptxGenJS();

    // Define slide master
    pptx.defineSlideMaster({
      title: masterSlide.name,
      background: masterSlide.background?.color
        ? { color: masterSlide.background.color }
        : masterSlide.background?.image
        ? { path: masterSlide.background.image }
        : { color: 'FFFFFF' },
    });

    // Create demo slide showing master configuration
    const slide = pptx.addSlide();
    slide.addText(`Master Slide: ${masterSlide.name}`, {
      x: 1,
      y: 1,
      fontSize: 32,
      bold: true,
    });

    const configInfo = [
      `Background: ${masterSlide.background?.color || masterSlide.background?.image || 'Default'}`,
      `Placeholders: ${masterSlide.placeholders?.length || 0}`,
      `Title Font: ${masterSlide.fonts?.title || 'Default'}`,
      `Body Font: ${masterSlide.fonts?.body || 'Default'}`,
      `Accent Colors: ${Object.keys(masterSlide.colors || {}).length}`,
    ];

    slide.addText(configInfo.join('\n'), {
      x: 1,
      y: 2.5,
      fontSize: 16,
      color: '666666',
    });

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addHyperlinks(
    filename: string,
    slideNumber: number,
    links: Array<{
      text: string;
      url?: string;
      slide?: number;
      tooltip?: string;
    }>
  ): Promise<Buffer> {
    const pptx = new PptxGenJS();

    const slide = pptx.addSlide();
    slide.addText(`Slide ${slideNumber} - Hyperlinks`, {
      x: 0.5,
      y: 0.5,
      fontSize: 32,
      bold: true,
    });

    // Add hyperlinks
    let yPos = 1.5;
    for (const link of links) {
      if (link.url) {
        slide.addText(link.text, {
          x: 1,
          y: yPos,
          fontSize: 18,
          color: '0066CC',
          underline: { style: 'sng' },
          hyperlink: { url: link.url, tooltip: link.tooltip },
        });
      } else if (link.slide) {
        slide.addText(link.text, {
          x: 1,
          y: yPos,
          fontSize: 18,
          color: '0066CC',
          underline: { style: 'sng' },
          hyperlink: { slide: link.slide, tooltip: link.tooltip },
        });
      }
      yPos += 0.5;
    }

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addSections(
    filename: string,
    sections: Array<{
      name: string;
      startSlide: number;
    }>
  ): Promise<Buffer> {
    const pptx = new PptxGenJS();

    // Note: PptxGenJS has limited section support
    // Create slides representing sections
    for (const section of sections) {
      const slide = pptx.addSlide();
      slide.addText(section.name, {
        x: 1,
        y: 2.5,
        w: '80%',
        fontSize: 44,
        bold: true,
        align: 'center',
        color: '0088CC',
      });

      slide.addText(`Section starts at slide ${section.startSlide}`, {
        x: 1,
        y: 4,
        w: '80%',
        fontSize: 18,
        align: 'center',
        color: '666666',
        italic: true,
      });
    }

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addMorphTransition(
    filename: string,
    fromSlide: number,
    toSlide: number,
    duration?: number
  ): Promise<Buffer> {
    const pptx = new PptxGenJS();

    // Slide 1
    const slide1 = pptx.addSlide();
    slide1.addText('Morph Transition - Slide 1', {
      x: 1,
      y: 1,
      fontSize: 32,
      bold: true,
    });
    slide1.addShape(pptx.ShapeType.rect, {
      x: 2,
      y: 3,
      w: 2,
      h: 1.5,
      fill: { color: '0088CC' },
    });

    // Slide 2
    const slide2 = pptx.addSlide();
    slide2.addText('Morph Transition - Slide 2', {
      x: 1,
      y: 1,
      fontSize: 32,
      bold: true,
    });
    slide2.addShape(pptx.ShapeType.rect, {
      x: 5,
      y: 3,
      w: 2,
      h: 1.5,
      fill: { color: 'CC0088' },
    });

    slide2.addText(
      `Note: Morph transition from slide ${fromSlide} to ${toSlide}\nDuration: ${duration || 1000}ms\n\nRequires PowerPoint 2016+ to apply actual morph effect.`,
      {
        x: 1,
        y: 5,
        fontSize: 14,
        color: '666666',
        italic: true,
      }
    );

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }

  async addActionButtons(
    filename: string,
    slideNumber: number,
    buttons: Array<{
      text: string;
      action: 'nextSlide' | 'previousSlide' | 'firstSlide' | 'lastSlide' | 'endShow' | 'customSlide';
      targetSlide?: number;
      x: number;
      y: number;
      w?: number;
      h?: number;
    }>
  ): Promise<Buffer> {
    const pptx = new PptxGenJS();

    const slide = pptx.addSlide();
    slide.addText(`Slide ${slideNumber} - Action Buttons`, {
      x: 0.5,
      y: 0.5,
      fontSize: 32,
      bold: true,
    });

    // Add action buttons
    for (const button of buttons) {
      const buttonWidth = button.w || 2;
      const buttonHeight = button.h || 0.75;

      // Draw button shape
      slide.addShape(pptx.ShapeType.rect, {
        x: button.x,
        y: button.y,
        w: buttonWidth,
        h: buttonHeight,
        fill: { color: '0088CC' },
        line: { color: '003366', width: 2 },
      });

      // Add button text
      slide.addText(button.text, {
        x: button.x,
        y: button.y,
        w: buttonWidth,
        h: buttonHeight,
        fontSize: 16,
        color: 'FFFFFF',
        bold: true,
        align: 'center',
        valign: 'middle',
      });

      // Add action metadata as small text below
      const actionText = button.action === 'customSlide'
        ? `Go to slide ${button.targetSlide}`
        : button.action.replace(/([A-Z])/g, ' $1').trim();

      slide.addText(actionText, {
        x: button.x,
        y: button.y + buttonHeight + 0.1,
        w: buttonWidth,
        fontSize: 10,
        color: '666666',
        align: 'center',
        italic: true,
      });
    }

    slide.addText(
      'Note: Interactive action buttons require PowerPoint to configure hyperlinks.',
      {
        x: 0.5,
        y: 6.5,
        fontSize: 12,
        color: '999999',
        italic: true,
      }
    );

    const buffer = await pptx.write({ outputType: 'arraybuffer' });
    return Buffer.from(buffer as ArrayBuffer);
  }
}
