var pptxhelper=pptxhelper||{};pptxhelper.Presentation=class{constructor(){this.pptxGen=new PptxGenJS,this.title="",this.width=10,this.height=5.625,this.slides=[]}addSlide(e,t){var i=new pptxhelper.Slide(this,e,t);return this.slides.push(i),$(i).on("slide.ready",()=>{this.checkSlideReady()}),i}download(e){this.isReady()?(void 0===e&&(e=this.title),this.pptxGen.save(e)):$(this).one("pptx.ready",()=>{this.download(e)})}setLayout(e,t,i){void 0===i&&(i="LAYOUT_"+parseInt(e)+"x"+parseInt(t)),this.width=e,this.height=t,this.pptxGen.setLayout({name:i,width:e,height:t})}setCompany(e){return this.pptxGen.setCompany(e),this}setTitle(e){return this.title=e,this.pptxGen.setTitle(e),this}setSubject(e){return this.pptxGen.setSubject(e),this}isReady(){var e=!0;return $(this.slides).each(function(){if(!this.isReady())return e=!1,!1}),e}checkSlideReady(){var e=!0;$(this.slides).each(function(){if(!this.isReady())return e=!1,!1}),e&&$(this).trigger("pptx.ready")}},pptxhelper.Slide=class{constructor(e,t,i){this.presentation=e,this.pendingData=0,this.pendingElements=[],this.slideGen=this.presentation.pptxGen.addNewSlide(),void 0!==t?(this.templateData=i,this.loadTemplate(t)):this.templateData={}}isReady(){return 0==this.pendingData}setBackground(e){return this.slideGen.bkgd=e,this}setTextColor(e){this.slideGen.color=e}addImage(e,t,i,a,s,n){if(this.pendingData>0)this.pendingElements.push({type:"image",x:e,y:t,width:i,height:a,source:s,hyperlink:n});else{if(s instanceof Image)r=this.imageToDataURL(s);else if(void 0!==s.tagName&&"canvas"==s.tagName.toLowerCase())r=s.toDataURL("image/png");else{if("string"!=typeof s||!s.startsWith("data:image/png;"))return++this.pendingData,void this.imageURLToDataURL(s,s=>{--this.pendingData,this.addImage(e,t,i,a,s),this.processPendingElements(),0==this.pendingData&&$(this).trigger("slide.ready")});var r=s}this.slideGen.addImage({x:e,y:t,w:i,h:a,data:r,hyperlink:this.createHyperlinkRef(n)})}}addText(e,t,i,a,s,n){this.pendingData>0&&this.pendingElements.push({type:"text",x:e,y:t,text:i,hyperlink:a,fontface:s,fontsize:n}),this.slideGen.addText(this.textPlaceholder(i),{x:e,y:t,font_face:s||"Arial",font_size:n||20,hyperlink:this.createHyperlinkRef(a)})}processPendingElements(){this.loadTemplate(this.pendingElements),this.pendingElements.length=0}createHyperlinkRef(e){return void 0!==e?{url:this.textPlaceholder(e)}:void 0}loadTemplate(e){var t=this;$(e).each(function(){var e=this;if("text"==e.type)t.addText(e.x,e.y,e.text,e.hyperlink,e.fontface,e.fontsize);else{if("image"!=e.type)throw new Error("Unknown template element type "+e.type);t.addImage(e.x,e.y,e.width,e.height,e.source,e.hyperlink)}})}imageToDataURL(e){if(void 0!==e.dataURL)return e.dataURL;var t=document.createElement("canvas");return t.width=e.naturalWidth,t.height=e.naturalHeight,t.getContext("2d").drawImage(e,0,0),e.dataURL=t.toDataURL("image/png"),e.dataURL}textPlaceholder(e){for(var t=-1;;){var i=e.indexOf("%(",t+1);if(-1==i)break;var a=e.indexOf(")",i),s=e.substring(i+2,a),n=this.templateData[s];e=e.replace("%("+s+")",n),t=i+n.length}return e}imageURLToDataURL(e,t){var i=new Image;$(i).on("load",()=>{t(this.imageToDataURL(i))}),i.src=e}};var xlsxhelper={};xlsxhelper.Workbook=class{constructor(){this.SheetNames=[],this.Sheets={}}addSheet(e){-1==this.SheetNames.indexOf(e.name)&&this.SheetNames.push(e.name),this.Sheets[e.name]=e}createFile(e,t,i){var a=e,s={bookType:"xls"==e?"biff2":e,bookSST:!1,type:"binary"},n=t?t+"."+a:null,r=XLSX.write(this,s);if("csv"!=e){for(var l=new ArrayBuffer(r.length),h=new Uint8Array(l),d=0;d<r.length;++d)h[d]=255&r.charCodeAt(d);r=l}return i&&i(n,r,"csv"!=e),r}},xlsxhelper.Sheet=class{constructor(e,t,i){this.name=e,this._disableRangeCalculation=!0;for(var a=0;a<t.length;++a)for(var s=t[a],n=0;n<s.length;++n){var r=s[n];null!=r&&this.setCell(n,a,r,i)}this._disableRangeCalculation=!1,this._updateRange()}static fromCSV(e,t,i){void 0===i&&(i={});var a=i.stringDelimiter||'"',s=i.cellDelimiter||",",n=i.lineDelimiter||"\n",r=i.keepDateWithOffset||!0,l=e.split(n),h=[];return $(l).each(function(){var e=this.toString(),t=[];let i=0;for(;i<e.length;){var n=e[i];if(++i,n==a){let s=e.indexOf(a,i);for(;;){if(-1==s)throw new Error('No closing quote found; line: "'+e+'"');if("\\"!=e[s-1])break;s=e.indexOf(a,s+1)}t.push(e.substring(i,s)),i=s+1}else if(n==s)t.push(null);else{let a=e.indexOf(s,i);-1==a&&(a=e.length),t.push(e.substring(i-1,a)),i=a}if(i<e.length&&e[i]!=s)throw new Error('Unknown cell delimiter: "'+e[i]+'"');++i}h.push(t)}),new xlsxhelper.Sheet(t,h,{keepDateWithOffset:r})}getCell(e,t){return this[XLSX.utils.encode_cell({c:e,r:t})]}setCell(e,t,i,a){void 0===a&&(a={});var s=a.keepDateWithOffset||!0;const n=/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.\d{3})?([+-]\d{2}:?\d{2}|Z)?$/;var r,l=XLSX.utils.encode_cell({c:e,r:t});if(void 0===i)delete this[l];else{if(i instanceof Date)if(0!=i.getTimezoneOffset()&&s)var h="s",d=i.toLocaleString();else var h="d",d=i.toISOString();else if(null!==(r=n.exec(i)))if(s&&"Z"!=r[8]&&"+00:00"!=r[8]&&"-00:00"!=r[8]&&"+0000"!=r[8]&&"-0000"!=r[8])var h="s",d=i;else{var h="d",d=i;i=new Date(i)}else if("boolean"==typeof i)var h="b",d=void 0;else if("number"==typeof i||parseFloat(i)==i){h="n";i=parseFloat(i);d=void 0}else var h="s",d=i;var o={v:i,t:h,w:d};this[l]=o}this._updateRange()}_updateRange(){if(!this._disableRangeCalculation){var e=null,t=null,i=null,a=null;for(let n in this)if(void 0!=this[n].t){var s=XLSX.utils.decode_cell(n);NaN!=s.c&&NaN!=s.r&&(null===e?(t=e=s.c,a=i=s.r):(s.c<e?e=s.c:s.c>t&&(t=s.c),s.r<i?i=s.r:s.r>a&&(a=s.r)))}null===e&&(e=0,t=0,i=0,a=0);var n={s:{c:e,r:i},e:{c:t,r:a}};this["!ref"]=XLSX.utils.encode_range(n)}}};
