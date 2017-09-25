/** Helper to use PptxGenJS */
/* @license
* Copyright 2017 Gabriel Paul 'Cley Faye' Risterucci
*
* Permission is hereby granted, free of charge, to any person obtaining a copy 
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is 
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in all
* copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
* SOFTWARE.
*/
var pptxhelper = pptxhelper || {};
/** A single presentation. */
pptxhelper.Presentation = class Presentation {
    constructor() {
        this.pptxGen = new PptxGenJS();
        this.title = '';
        // Default layout size
        this.width = 10;
        this.height = 5.625;
        this.slides = [];
    }
    /** Add a new slide to the presentation.
    *
    * If template is undefined, the new slide is empty. Otherwise template and
    * data are parsed according to the Slide.loadTemplate() doc.
    */
    addSlide(template, data) {
        var slide = new pptxhelper.Slide(this, template, data);
        this.slides.push(slide);
        $(slide).on('slide.ready', () => {
            this.checkSlideReady();
        });
        return slide;
    }
    /** Create a file from the presentation. */
    download(filename) {
        if (!this.isReady()) {
            $(this).one('pptx.ready', () => {
                this.download(filename);
            });
            return;
        }
        if (filename === undefined) {
            filename = this.title;
        }
        this.pptxGen.save(filename);
    }
    setLayout(width, height, name) {
        if (name === undefined) {
            name = 'LAYOUT_' + parseInt(width) + 'x' + parseInt(height);
        }
        this.width = width;
        this.height = height;
        this.pptxGen.setLayout({
            name: name,
            width: width,
            height: height});
    }
    setCompany(companyName) {
        this.pptxGen.setCompany(companyName);
        return this;
    }
    setTitle(title) {
        this.title = title;
        this.pptxGen.setTitle(title);
        return this;
    }
    setSubject(subject) {
        this.pptxGen.setSubject(subject);
        return this;
    }
    isReady() {
        var allSlideReady = true;
        $(this.slides).each(function() {
            if (!this.isReady()) {
                allSlideReady = false;
                return false;
            }
        });
        return allSlideReady;
    }
    checkSlideReady() {
        var allReady = true;
        $(this.slides).each(function() {
            if (!this.isReady()) {
                allReady = false;
                return false;
            }
        });
        if (allReady) {
            $(this).trigger('pptx.ready');
        }
    }
};
/** A single slide in a PPTX presentation.
*
* Since some content can be downloaded asynchronously, a slide will trigger a
* "slide.ready" event when everything's downloaded.
* This event will propagate to the parent PPTX instance that will trigger a
* "pptx.ready" event on itself.
*/
pptxhelper.Slide = class Slide {
    constructor(presentation, template, data) {
        this.presentation = presentation;
        this.pendingData = 0;
        this.slideGen = this.presentation.pptxGen.addNewSlide();
        if (template !== undefined) {
            this.templateData = data;
            this.loadTemplate(template);
        } else {
            this.templateData = {};
        }
    }
    /** Determine if the slide is ready (all images are downloaded) */
    isReady() {
        return this.pendingData == 0;
    }
    setBackground(color) {
        this.slideGen.bkgd = color;
        return this;
    }
    setTextColor(color) {
        this.slideGen.color = color;
    }
    /** Add an image to the slide.
    *
    * Parameters
    * ----------
    * image : Image | Canvas | string
    *     The image to add. If an URL is provided, the image is downloaded
    *     first.
    * hyperlink : string (optional)
    *     Add a clickable link
    */
    addImage(x, y, width, height, image, hyperlink) {
        if (image instanceof Image) {
            var dataURL = this.imageToDataURL(image);
        } else if (image.tagName !== undefined 
                   && image.tagName.toLowerCase() == 'canvas') {
            var dataURL = image.toDataURL('image/png');
        } else if (typeof(image) == 'string'
                   && image.startsWith('data:image/png;')) {
            var dataURL = image;
        } else {
            ++this.pendingData;
            this.imageURLToDataURL(image, (dataURL) => {
                --this.pendingData;
                this.addImage(x, y, width, height, dataURL);
                if (this.pendingData == 0) {
                    $(this).trigger('slide.ready');
                }
            });
            return;
        }
        this.slideGen.addImage({
            x: x,
            y: y,
            w: width,
            h: height,
            data: dataURL,
            hyperlink: this.createHyperlinkRef(hyperlink)});
    }
    /** Adds text to the slide.
    * 
    * Parameters
    * ----------
    * x, y : number
    *     Text position (from top-left corner)
    * text : string
    *     Text to display
    * hyperlink : string (optional)
    *     Add a clickable link
    * font_face : string (optional)
    *     Name of the font to use. Default to 'Arial'
    * font_size : number (optional)
    *     Size of the font (default to 20)
    */
    addText(x, y, text, hyperlink, font_face, font_size) {
        this.slideGen.addtext(text, {
            x: x,
            y: y,
            font_face: font_face || 'Arial',
            font_size: font_size || 20,
            hyperlink: this.createHyperlinkRef(hyperlink)});
    }
    createHyperlinkRef(hyperlink) {
        if (hyperlink !== undefined) {
            return { url: this.textPlaceholder(hyperlink) };
        } else {
            return undefined;
        }
    }
    /** Load a slide definition from a template.
    *
    * The template is an array where each element can be either a text or an
    * image definition.
    *
    * All objects definitions share the following properties:
    * 
    * - 'type': Type of the element ('text' or 'image')
    * - 'x': X offset (from left side)
    * - 'y': Y offset (from top side)
    * - hyperlink (optional): transform the text to an hyperlink
    *
    * Text definitions have the following properties:
    *
    * - type=='text'
    * - text: Text to display
    * - fontface (optional): font to display the text with
    * - fontsize (optional): size of the font
    *
    * Image definitions have the following properties:
    * - type=='image'
    * - width: image width
    * - height: image height
    * - source: either an Image object, a canvas, a PNG DataURL, or an image
    *   URL.
    *
    * Displayed text can include some markers to identify text placeholders
    * provided by the data parameter.
    * Such placeholder are marked using %(placeholder_name). They will look into
    * data properties to find one with a matching name, and replace the
    * placeholder by the property value.
    * This also works for hyperlinks.
    */
    loadTemplate(template) {
        for (var element in template) {
            if (element.type == 'text') {
                this.addText(element.x,
                             element.y,
                             element.text,
                             element.hyperlink,
                             element.fontface,
                             element.fontsize);
            } else if (element.type == 'image') {
                this.addImage(element.x,
                              element.y,
                              element.width,
                              element.height,
                              element.source,
                              element.hyperlink);
            } else {
                throw  new Error('Unknown template element type ' 
                                 + element.type);
            }
        }
    }
    /** Convert an image to a data URL.
    *
    * This is synchronous.
    */
    imageToDataURL(image) {
        if (image.dataURL !== undefined) {
            return image.dataURL;
        }
        var canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        canvas.getContext('2d').drawImage(image, 0, 0);
        image.dataURL = canvas.toDataURL('image/png');
        return image.dataURL;
    }
    /** Convert a (potential) placeholder text to its final value. */
    textPlaceholder(text) {
        var lastIndex = -1;
        do {
            var placeholderStartIndex = text.indexOf('%(', lastIndex + 1);
            if (placeholderStartIndex == -1) {
                break;
            }
            var placeholderEndIndex = text.indexOf(')', placeholderStartIndex);
            var placeholderName = text.substring(placeholderStartIndex + 2,
                                                 placeholderEndIndex);
            var replacement = this.templateData[placeholderName];
            text.replace('%(' + placeholderName + ')', replacement);
            lastIndex = placeholderStartIndex + replacement.length;
        } while (true);
        return text;
    }
    /** Download an image and convert it to a data URL.
    *
    * This is asynchronous.
    *
    * Parameters
    * ----------
    * cb : function
    *     Callback that will receive the dataurl as a parameter.
    */
    imageURLToDataURL(url, cb) {
        var img = new Image();
        $(img).on('load', () => {
            cb(this.imageToDataURL(img));
        });
        img.src = url;
    }
};
