/* AK TOOLS — faithful converters (loaded last, overrides the basic handlers)
   PPTX -> PDF : every slide rendered as it looks (images + text in place)
   PDF -> DOCX : page text + the PDF's own images, kept in the Word file */
(function(){
"use strict";
var $=function(id){return document.getElementById(id)};
var toastEl=$("toast");
function toast(msg,err){ if(!toastEl)return; toastEl.textContent=msg; toastEl.className="toast show"+(err?" err":""); clearTimeout(toast._t); toast._t=setTimeout(function(){toastEl.classList.remove("show")},2600); }
function busy(btn,on,label){ if(!btn)return; if(on){btn.dataset.orig=btn.textContent;btn.disabled=true;if(label)btn.textContent=label;}else{btn.disabled=false;if(btn.dataset.orig)btn.textContent=btn.dataset.orig;} }
function prog(id,pct,text){ var P=$(id); if(!P)return; P.classList.add("on"); var f=P.querySelector(".fill"),t=P.querySelector(".pt"); if(f)f.style.width=Math.round(pct)+"%"; if(t&&text!=null)t.textContent=text; }
function progDone(id){ setTimeout(function(){ var P=$(id); if(!P)return; P.classList.remove("on"); var f=P.querySelector(".fill"); if(f)f.style.width="0%"; },700); }
function setDl(zoneId,blob,name){
  var z=$(zoneId); if(!z)return; z.innerHTML=""; z.classList.add("show");
  var url=URL.createObjectURL(blob);
  var a=document.createElement("a"); a.className="dl"; a.href=url; a.download=name; a.textContent="\u2B07 Download "+name;
  var inp=document.createElement("input"); inp.className="dlname"; inp.value=name; inp.setAttribute("aria-label","File name");
  inp.addEventListener("input",function(){ var v=inp.value.trim(); var ext=name.slice(name.lastIndexOf(".")); if(v&&v.toLowerCase().slice(-ext.length)!==ext)v+=ext; a.download=v||name; a.textContent="\u2B07 Download "+(v||name); });
  z.appendChild(inp); z.appendChild(a);
  var note=document.createElement("div"); note.className="note"; note.textContent="Tap the name above to rename before saving."; z.appendChild(note);
}
if(window.pdfjsLib&&!window.pdfjsLib.GlobalWorkerOptions.workerSrc){
  window.pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}
/* keep our own copy of the chosen files (browse, drag-drop and paste all land here) */
var myFiles={};
function capture(sectionId,key){
  var sec=$(sectionId); if(!sec)return;
  var drop=sec.querySelector(".drop"), inp=sec.querySelector('input[type="file"]');
  if(inp)inp.addEventListener("change",function(){ if(inp.files&&inp.files[0])myFiles[key]=inp.files[0]; });
  if(drop)drop.addEventListener("drop",function(e){ var f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0]; if(f)myFiles[key]=f; });
}
capture("tool-ppt2pdf","pptx");
capture("tool-pdf2word","pdf");

/* ---------------- PPTX -> PDF (render slides as they look) ---------------- */
function parseXml(txt){ return new DOMParser().parseFromString(txt,"text/xml"); }
function relsMap(doc){ var m={},rs=doc.getElementsByTagName("Relationship"),i; for(i=0;i<rs.length;i++)m[rs[i].getAttribute("Id")]=rs[i].getAttribute("Target"); return m; }
function resolvePath(baseDir,target){
  var parts=(baseDir+target).split("/"),out=[],i,p;
  for(i=0;i<parts.length;i++){ p=parts[i]; if(p==="..")out.pop(); else if(p&&p!==".")out.push(p); }
  return out.join("/");
}
function srgbOf(el){ if(!el)return null; var c=el.getElementsByTagName("a:srgbClr")[0]; return c?"#"+c.getAttribute("val"):null; }
function loadImg(blob){ return new Promise(function(res,rej){ var u=URL.createObjectURL(blob); var im=new Image(); im.onload=function(){res({im:im,u:u})}; im.onerror=function(e){URL.revokeObjectURL(u);rej(e)}; im.src=u; }); }
function xfrm(el){
  var xf=el.getElementsByTagName("a:xfrm")[0]; if(!xf)return null;
  var off=xf.getElementsByTagName("a:off")[0], ext=xf.getElementsByTagName("a:ext")[0];
  if(!off||!ext)return null;
  return {x:+off.getAttribute("x")||0, y:+off.getAttribute("y")||0, w:+ext.getAttribute("cx")||0, h:+ext.getAttribute("cy")||0};
}
function drawTextShape(el,ctx,S,pxPerInch){
  var box=xfrm(el); if(!box||box.w<=0||box.h<=0)return;
  var bx=box.x*S, by=box.y*S, bw=box.w*S, bh=box.h*S;
  var paras=el.getElementsByTagName("a:p"); if(!paras.length)return;
  ctx.save(); ctx.beginPath(); ctx.rect(bx,by,bw,bh); ctx.clip();
  var cy=by+4;
  for(var i=0;i<paras.length;i++){
    var p=paras[i], pPr=p.getElementsByTagName("a:pPr")[0];
    var algn=pPr?pPr.getAttribute("algn"):null;
    var runs=p.getElementsByTagName("a:r"), text="", maxSz=1800, bold=false, color="#111111", j, r, rPr;
    for(j=0;j<runs.length;j++){
      r=runs[j]; var t=r.getElementsByTagName("a:t")[0]; if(t)text+=t.textContent;
      rPr=r.getElementsByTagName("a:rPr")[0];
      if(rPr){ var sz=+rPr.getAttribute("sz")||0; if(sz>maxSz)maxSz=sz; if(rPr.getAttribute("b")==="1")bold=true; var col=srgbOf(rPr); if(col)color=col; }
    }
    if(!text){ cy+=10; continue; }
    var fontPx=(maxSz/100)*(pxPerInch/72);
    ctx.font=(bold?"700 ":"400 ")+fontPx+"px Arial, sans-serif";
    ctx.fillStyle=color; ctx.textBaseline="top";
    var words=text.split(/\s+/), lines=[], line="";
    for(j=0;j<words.length;j++){ var test=line?line+" "+words[j]:words[j]; if(ctx.measureText(test).width>bw-8&&line){lines.push(line);line=words[j];}else line=test; }
    if(line)lines.push(line);
    var lh=fontPx*1.25;
    for(j=0;j<lines.length;j++){
      var lx=bx+4; ctx.textAlign="left";
      if(algn==="ctr"){ctx.textAlign="center";lx=bx+bw/2;}
      else if(algn==="r"){ctx.textAlign="right";lx=bx+bw-4;}
      ctx.fillText(lines[j],lx,cy); cy+=lh;
      if(cy>by+bh)break;
    }
    cy+=fontPx*0.25;
    if(cy>by+bh)break;
  }
  ctx.restore();
}
function drawShapes(parent,zip,slideDir,rels,ctx,S,pxPerInch,done){
  var kids=[],i; for(i=0;i<parent.childNodes.length;i++){ if(parent.childNodes[i].nodeType===1)kids.push(parent.childNodes[i]); }
  var chain=Promise.resolve();
  kids.forEach(function(el){
    chain=chain.then(function(){
      var tag=el.tagName;
      if(tag==="p:pic"){
        var blip=el.getElementsByTagName("a:blip")[0]; if(!blip)return;
        var rid=blip.getAttribute("r:embed"), target=rels[rid]; if(!target)return;
        var box=xfrm(el); if(!box)return;
        var zf=zip.file(resolvePath(slideDir,target)); if(!zf)return;
        return zf.async("blob").then(function(b){ return loadImg(b); }).then(function(o){
          ctx.drawImage(o.im, box.x*S, box.y*S, box.w*S, box.h*S); URL.revokeObjectURL(o.u);
        }).catch(function(){});
      }
      if(tag==="p:grpSp"){ return drawShapes(el,zip,slideDir,rels,ctx,S,pxPerInch); }
      if(tag==="p:sp"||tag==="p:cxnSp"||tag==="p:graphicFrame"){ drawTextShape(el,ctx,S,pxPerInch); }
    });
  });
  return chain;
}
var ppBtn=$("ppConvert");
if(ppBtn)ppBtn.onclick=function(){
  var f=myFiles.pptx; if(!f)return toast("Add a .pptx file first",1);
  busy(ppBtn,true,"Converting\u2026");
  (async function(){
    try{
      prog("ppProg",4,"Reading presentation\u2026");
      var zip=await JSZip.loadAsync(f);
      var presDoc=parseXml(await zip.file("ppt/presentation.xml").async("text"));
      var presRels=relsMap(parseXml(await zip.file("ppt/_rels/presentation.xml.rels").async("text")));
      var szEl=presDoc.getElementsByTagName("p:sldSz")[0];
      var EMUw=szEl?+szEl.getAttribute("cx")||12192000:12192000;
      var EMUh=szEl?+szEl.getAttribute("cy")||6858000:6858000;
      var mmW=EMUw/36000, mmH=EMUh/36000;
      var Wpx=1600, S=Wpx/EMUw, Hpx=Math.round(EMUh*S), pxPerInch=Wpx/(EMUw/914400);
      var idEls=presDoc.getElementsByTagName("p:sldId"), targets=[], i;
      for(i=0;i<idEls.length;i++){ var t=presRels[idEls[i].getAttribute("r:id")]; if(t)targets.push(t); }
      if(!targets.length)throw new Error("no slides");
      var pdf=new window.jspdf.jsPDF({orientation:mmW>=mmH?"landscape":"portrait", unit:"mm", format:[mmW,mmH], compress:true});
      for(i=0;i<targets.length;i++){
        prog("ppProg", 8+Math.round((i/targets.length)*86), "Slide "+(i+1)+" of "+targets.length+"\u2026");
        var slidePath=resolvePath("ppt/", targets[i]);
        var slideDir=slidePath.slice(0, slidePath.lastIndexOf("/")+1);
        var sDoc=parseXml(await zip.file(slidePath).async("text"));
        var sRels={}; var rf=zip.file(slideDir+"_rels/"+slidePath.split("/").pop()+".rels");
        if(rf)sRels=relsMap(parseXml(await rf.async("text")));
        var cv=document.createElement("canvas"); cv.width=Wpx; cv.height=Hpx;
        var ctx=cv.getContext("2d");
        var bgEl=sDoc.getElementsByTagName("p:bg")[0]; var bg=bgEl?srgbOf(bgEl):null;
        ctx.fillStyle=bg||"#ffffff"; ctx.fillRect(0,0,Wpx,Hpx);
        var tree=sDoc.getElementsByTagName("p:spTree")[0];
        if(tree)await drawShapes(tree,zip,slideDir,sRels,ctx,S,pxPerInch);
        if(i>0)pdf.addPage([mmW,mmH], mmW>=mmH?"landscape":"portrait");
        pdf.addImage(cv.toDataURL("image/jpeg",0.88),"JPEG",0,0,mmW,mmH);
      }
      var base=f.name.replace(/\.pptx$/i,"");
      setDl("ppDl", pdf.output("blob"), "aktools-"+base+".pdf");
      prog("ppProg",100,"Done \u2713"); toast("PDF ready \u2014 slides kept as-is \u2713");
    }catch(err){ toast("Conversion failed \u2014 is it a valid .pptx?",1); }
    busy(ppBtn,false); progDone("ppProg");
  })();
};

/* ---------------- PDF -> DOCX (text + images kept) ---------------- */
function pageImages(pg){
  var OPS=window.pdfjsLib.OPS, out=[], seen={};
  return pg.getOperatorList().then(function(ops){
    var chain=Promise.resolve();
    ops.fnArray.forEach(function(fn,idx){
      if(fn!==OPS.paintImageXObject&&fn!==OPS.paintInlineImageXObject)return;
      var name=ops.argsArray[idx][0]; if(seen[name])return; seen[name]=1;
      chain=chain.then(function(){
        return new Promise(function(res){
          var done=false; function cb(v){ if(!done){done=true;res(v||null);} }
          try{ var r=pg.objs.get(name,cb); if(r)cb(r); }catch(e){ res(null); }
          setTimeout(function(){cb(null)},2000);
        }).then(function(img){
          if(!img)return;
          try{
            var cv=document.createElement("canvas");
            if(img.bitmap){ cv.width=img.width; cv.height=img.height; cv.getContext("2d").drawImage(img.bitmap,0,0); }
            else if(typeof HTMLCanvasElement!=="undefined"&&(img instanceof HTMLCanvasElement||img instanceof HTMLImageElement)){ cv.width=img.width; cv.height=img.height; cv.getContext("2d").drawImage(img,0,0); }
            else if(img.data&&img.width&&img.height){
              cv.width=img.width; cv.height=img.height;
              var ctx=cv.getContext("2d"), id=ctx.createImageData(img.width,img.height), kind=img.kind, s=0, p;
              if(kind===3){ id.data.set(img.data); }
              else{ var ch=kind===1?1:3; for(p=0;p<id.data.length;p+=4){ var v=img.data[s]; id.data[p]=v; id.data[p+1]=kind===1?v:img.data[s+1]; id.data[p+2]=kind===1?v:img.data[s+2]; id.data[p+3]=255; s+=ch; } }
              ctx.putImageData(id,0,0);
            } else return;
            if(cv.width>=8&&cv.height>=8&&out.length<14)out.push(cv);
          }catch(e){}
        });
      });
    });
    return chain;
  }).then(function(){ return out; }).catch(function(){ return out; });
}
var pwBtn=$("p2wConvert");
if(pwBtn)pwBtn.onclick=function(){
  var f=myFiles.pdf; if(!f)return toast("Add a PDF file first",1);
  busy(pwBtn,true,"Converting\u2026");
  (async function(){
    try{
      prog("p2wProg",4,"Reading PDF\u2026");
      var buf=await f.arrayBuffer();
      var pdf=await window.pdfjsLib.getDocument({data:buf}).promise;
      var D=window.docx, children=[], i;
      for(i=1;i<=pdf.numPages;i++){
        var pg=await pdf.getPage(i);
        var tc=await pg.getTextContent();
        var lines=[], cur="", lastY=null;
        tc.items.forEach(function(it){ var y=Math.round(it.transform[5]); if(lastY!==null&&y!==lastY){lines.push(cur);cur="";} cur+=it.str; lastY=y; if(it.hasEOL){lines.push(cur);cur="";} });
        if(cur)lines.push(cur);
        if(!lines.length)lines=[""];
        lines.forEach(function(ln,idx){
          children.push(new D.Paragraph({ pageBreakBefore: i>1&&idx===0, children:[new D.TextRun({text:ln||" ", size:24})] }));
        });
        var cvs=await pageImages(pg), k;
        for(k=0;k<cvs.length;k++){
          var cv=cvs[k];
          var blob=await new Promise(function(r){ cv.toBlob(r,"image/png"); });
          if(!blob)continue;
          var data=new Uint8Array(await blob.arrayBuffer());
          var sc=Math.min(1, 620/cv.width);
          children.push(new D.Paragraph({ children:[ new D.ImageRun({ data:data, transformation:{ width:Math.round(cv.width*sc), height:Math.round(cv.height*sc) } }) ] }));
        }
        prog("p2wProg", 8+Math.round((i/pdf.numPages)*88), "Page "+i+" of "+pdf.numPages+"\u2026");
      }
      var outBlob=await D.Packer.toBlob(new D.Document({ sections:[{ children:children }] }));
      var base=f.name.replace(/\.pdf$/i,"");
      setDl("p2wDl", outBlob, "aktools-"+base+".docx");
      prog("p2wProg",100,"Done \u2713"); toast("Word file ready \u2014 text + images \u2713");
    }catch(err){ toast("Conversion failed \u2014 is it a valid PDF?",1); }
    busy(pwBtn,false); progDone("p2wProg");
  })();
};
})();
