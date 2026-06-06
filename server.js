const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const admin = require('firebase-admin');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        ImageRun, AlignmentType, WidthType, BorderStyle, ShadingType,
        VerticalAlign, Header } = require('docx');

// ── Firebase init ─────────────────────────────────────────
const serviceAccount = require('./firebase-key.json');
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});
const db = admin.firestore();

const app = express();
app.use(cors());
app.use(express.json({ limit: '80mb' }));
app.use(express.static(__dirname));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'informe_clima_app.html')));

// ── Design tokens ─────────────────────────────────────────
const BL  = 'DEEAF6';
const BL2 = 'D9E2F3';
const GH  = 'D9D9D9';
const BC  = '7B7B7B';
const TW  = 9869;

const thinB = (color=BC) => ({ style: BorderStyle.SINGLE, size: 4, color });
const allThin = { top: thinB(), bottom: thinB(), left: thinB(), right: thinB() };

const mkPara = (text, opts={}) => new Paragraph({
  alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
  spacing: { before: 0, after: 0 },
  children: [new TextRun({
    text: text || '', bold: opts.bold||false, italics: opts.italics||false,
    size: opts.sz || 18, font: 'Calibri', color: opts.color||'000000'
  })]
});

const LC = (text, w, span=1, fill=BL) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  ...(span > 1 ? { columnSpan: span } : {}),
  borders: allThin,
  shading: { fill, type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 40, bottom: 40, left: 70, right: 40 },
  children: [mkPara(text, { bold: true, sz: 16 })]
});

const VC = (text, w, span=1) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  ...(span > 1 ? { columnSpan: span } : {}),
  borders: allThin,
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 40, bottom: 40, left: 70, right: 40 },
  children: [mkPara(text || '', { sz: 16 })]
});

const HC = (text, w, span=1, rowSpan=1) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  ...(span > 1 ? { columnSpan: span } : {}),
  ...(rowSpan > 1 ? { rowSpan } : {}),
  borders: allThin,
  shading: { fill: BL2, type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 30, bottom: 30, left: 30, right: 30 },
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: text||'', bold: true, size: 14, font: 'Calibri' })]
  })]
});

const secRow = (text) => new TableRow({
  height: { value: 300 },
  children: [new TableCell({
    width: { size: TW, type: WidthType.DXA }, columnSpan: 14,
    borders: allThin,
    shading: { fill: GH, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 100, right: 40 },
    children: [mkPara(text, { bold: true, sz: 18 })]
  })]
});

// ── Build DOCX ─────────────────────────────────────────────
async function buildDocx(d) {
  const v = s => (s||'').toString().trim() || 'N/A';

  let headerChildren = [];
  try {
    const logoData = fs.readFileSync(path.join(__dirname, 'logo.jpeg'));
    const headerTable = new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: [2563, 4625, 745, 1936],
      borders: {
        top: thinB('A5A5A5'), bottom: thinB('A5A5A5'),
        left: thinB('A5A5A5'), right: thinB('A5A5A5'),
        insideH: thinB('A5A5A5'), insideV: thinB('A5A5A5')
      },
      rows: [
        new TableRow({ height: { value: 410 }, children: [
          new TableCell({ width:{size:2563,type:WidthType.DXA}, rowSpan:2, borders:allThin,
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:40,right:40},
            children:[new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
              children:[new ImageRun({data:logoData,transformation:{width:110,height:30},type:'jpeg'})] })] }),
          new TableCell({ width:{size:4625,type:WidthType.DXA},
            borders:{top:thinB('A5A5A5'),bottom:thinB(BC),left:thinB('A5A5A5'),right:thinB('A5A5A5')},
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:70,right:40},
            children:[mkPara('CENTRALES CLIMA',{bold:true,sz:22,center:true})] }),
          new TableCell({ width:{size:745,type:WidthType.DXA},
            borders:{top:thinB('A5A5A5'),bottom:thinB('A5A5A5'),left:thinB('A5A5A5'),right:thinB(BC)},
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:40,right:40},
            children:[mkPara('COD.',{sz:16})] }),
          new TableCell({ width:{size:1936,type:WidthType.DXA},
            borders:{top:thinB('A5A5A5'),bottom:thinB('A5A5A5'),left:thinB(BC),right:thinB('A5A5A5')},
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:40,right:40},
            children:[mkPara(v(d.codInforme),{sz:16})] }),
        ]}),
        new TableRow({ height: { value: 320 }, children: [
          new TableCell({ width:{size:4625,type:WidthType.DXA},
            borders:{top:thinB(BC),bottom:thinB('A5A5A5'),left:thinB('A5A5A5'),right:thinB('A5A5A5')},
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:70,right:40},
            children:[new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
              children:[new TextRun({text:'INFORME CORRECTIVO CLIMA',bold:true,size:22,font:'Calibri'})] })] }),
          new TableCell({ width:{size:745,type:WidthType.DXA},
            borders:{top:thinB('A5A5A5'),bottom:thinB('A5A5A5'),left:thinB('A5A5A5'),right:thinB(BC)},
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:40,right:40},
            children:[mkPara('FECHA',{sz:16})] }),
          new TableCell({ width:{size:1936,type:WidthType.DXA},
            borders:{top:thinB('A5A5A5'),bottom:thinB('A5A5A5'),left:thinB(BC),right:thinB('A5A5A5')},
            verticalAlign:VerticalAlign.CENTER, margins:{top:20,bottom:20,left:40,right:40},
            children:[new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
              children:[new TextRun({text:v(d.fecha),size:16,font:'Calibri'})] })] }),
        ]})
      ]
    });
    headerChildren = [headerTable, new Paragraph({ spacing:{before:0,after:0}, children:[] })];
  } catch(e) {
    headerChildren = [new Paragraph({ children:[] })];
  }

  const scale = 9869 / 11231;
  const cw = [2827,517,528,782,269,1056,1050,262,787,527,523,790,261,1052];
  const scaled = cw.map(w => Math.round(w * scale));
  scaled[0] += TW - scaled.reduce((a,b)=>a+b,0);

  const row_ig = secRow('INFORMACION GENERAL');
  const row_sitio = new TableRow({ height:{value:280}, children:[
    LC('Nombre de Sitio',scaled[0]+scaled[1],2),
    VC(v(d.nombreSitio),scaled[2]+scaled[3]+scaled[4]+scaled[5],4),
    LC('Código de Sitio',scaled[6]+scaled[7]+scaled[8]+scaled[9],4),
    VC(v(d.codigoSitio),scaled[10]+scaled[11]+scaled[12]+scaled[13],4)
  ]});
  const row_dir = new TableRow({ height:{value:280}, children:[
    LC('Dirección',scaled[0]+scaled[1],2),
    VC(v(d.direccion),TW-(scaled[0]+scaled[1]),12)
  ]});

  const wl=scaled[0]+scaled[1], wi=scaled[2]+scaled[3], wte=scaled[4]+scaled[5],
        wti=scaled[6]+scaled[7], wr=scaled[8]+scaled[9],
        wot=scaled[10]+scaled[11], wotv=scaled[12]+scaled[13];

  const row_tk = new TableRow({ height:{value:280}, children:[
    new TableCell({ width:{size:wl,type:WidthType.DXA}, columnSpan:2, rowSpan:2,
      borders:allThin, shading:{fill:BL,type:ShadingType.CLEAR},
      verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:70,right:40},
      children:[mkPara('Números de Tickets',{bold:true,sz:16})] }),
    LC('Inc.',wi,2,BL2), LC('TE',wte,2,BL2), LC('TI',wti,2,BL2),
    LC('RED',wr,2,BL2), LC('Numero de OT',wot,2,BL2), VC('',wotv,2)
  ]});
  const row_tk2 = new TableRow({ height:{value:260}, children:[
    VC(v(d.ticketInc),wi,2), VC(v(d.ticketTE),wte,2),
    VC(v(d.ticketTI),wti,2), VC(v(d.ticketRED),wr,2),
    new TableCell({ width:{size:wot,type:WidthType.DXA}, columnSpan:2,
      borders:allThin, shading:{fill:BL2,type:ShadingType.CLEAR},
      verticalAlign:VerticalAlign.CENTER, margins:{top:40,bottom:40,left:40,right:40},
      children:[mkPara('')] }),
    VC(v(d.numOT),wotv,2)
  ]});
  const row_sala = new TableRow({ height:{value:280}, children:[
    LC('Sala',wl,2), VC(v(d.sala),wi+wte+wti+wr,8),
    LC('Fecha Ejecución',wot,2,BL2), VC(v(d.fecha),wotv,2)
  ]});
  const row_tec = new TableRow({ height:{value:280}, children:[
    LC('Técnico Ejecutante',wl,2), VC(v(d.tecnico),wi+wte,4),
    LC('Supervisor',wti+wr,4), VC(v(d.supervisor),wot+wotv,4)
  ]});

  const row_rs  = secRow('RESUMEN DE LA ACTIVIDAD');
  const row_rs2 = new TableRow({ height:{value:1600}, children:[
    new TableCell({ width:{size:TW,type:WidthType.DXA}, columnSpan:14,
      borders:allThin, margins:{top:60,bottom:60,left:100,right:100},
      children:[mkPara(v(d.resumen),{sz:16})] })
  ]});

  const row_eq = secRow('DATOS GENERALES DEL EQUIPAMIENTO');
  const eq1=Math.round(TW*0.23), eq2=Math.round(TW*0.17), eq3=Math.round(TW*0.19),
        eq4=Math.round(TW*0.22), eq5=TW-eq1-eq2-eq3-eq4;
  const row_eq_h = new TableRow({ height:{value:280}, children:[
    HC('Sala',eq1,1), HC('N° Equipo',eq2,4), HC('Tipo',eq3,2),
    HC('Marca',eq4,4), HC('Modelo / Serie',eq5,3)
  ]});
  const row_eq_d = new TableRow({ height:{value:280}, children:[
    VC(v(d.eqSala),eq1,1), VC(v(d.eqNumero),eq2,4),
    VC(v(d.eqTipo),eq3,2), VC(v(d.eqMarca),eq4,4), VC(v(d.eqModelo),eq5,3)
  ]});

  const row_med = secRow('MEDICIONES GENERALES');
  const mw = Math.floor(TW/14), mw2 = Math.floor((mw*3)/2);
  const row_med_h1 = new TableRow({ height:{value:260}, children:[
    HC('N° De Equipo',mw*2,1,2), HC('Consumo Compresor COMP 1',mw*3,4),
    HC('Consumo Evaporador',mw*2,2), HC('Consumo Condensador',mw*3,4),
    HC('Temperatura',TW-mw*2-mw*3-mw*2-mw*3,3)
  ]});
  const row_med_h2 = new TableRow({ height:{value:260}, children:[
    HC('V.Prom\n(Volt)',mw2,2), HC('Corriente\n(Amp)',mw*3-mw2,2),
    HC('V.Prom\n(Volt)',mw,1), HC('Corriente\n(Amp)',mw,1),
    HC('V.Prom\n(Volt)',mw2,2), HC('Corriente\n(Amp)',mw*3-mw2,2),
    HC('Inyección\n(°C)',mw,2), HC('Retorno\n(°C)',TW-mw*2-mw*3-mw*2-mw*3-mw,1)
  ]});
  const row_med_d = new TableRow({ height:{value:300}, children:[
    VC(v(d.eqNumero),mw*2,1),
    VC(v(d.m_cv),mw2,2), VC(v(d.m_ca),mw*3-mw2,2),
    VC(v(d.m_ev),mw,1), VC(v(d.m_ea),mw,1),
    VC(v(d.m_condv),mw2,2), VC(v(d.m_conda),mw*3-mw2,2),
    VC(v(d.m_tinj),mw,2), VC(v(d.m_tret),TW-mw*2-mw*3-mw*2-mw*3-mw,1)
  ]});

  const row_obs  = secRow('OBSERVACIONES Y RECOMENDACIONES');
  const row_obs2 = new TableRow({ height:{value:1400}, children:[
    new TableCell({ width:{size:TW,type:WidthType.DXA}, columnSpan:14,
      borders:allThin, margins:{top:60,bottom:60,left:100,right:100},
      children:[mkPara(v(d.observaciones),{sz:16})] })
  ]});

  const row_foto_hdr = secRow('REGISTRO FOTOGRAFICO');
  const photoRows = [];
  for (let r = 0; r < 6; r++) {
    const i1=r*2, i2=r*2+1;
    const p1=d.photos&&d.photos[i1], p2=d.photos&&d.photos[i2];
    if (!p1 && !p2) continue;
    const cells = [];
    for (const [idx, photoData] of [[i1,p1],[i2,p2]]) {
      let children = [];
      if (photoData) {
        const base64 = photoData.replace(/^data:image\/\w+;base64,/,'');
        const imgBuf = Buffer.from(base64,'base64');
        const ext = photoData.startsWith('data:image/png') ? 'png' : 'jpeg';
        const descText = (d.photoDescs&&d.photoDescs[idx]) ? d.photoDescs[idx] : `Fig. ${idx+1}`;
        children = [
          new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:0,after:0},
            children:[new ImageRun({data:imgBuf,transformation:{width:210,height:158},type:ext})] }),
          new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:40,after:0},
            children:[new TextRun({text:descText,italics:true,bold:true,size:14,font:'Calibri'})] })
        ];
      } else {
        children = [
          new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun({text:'',size:14})] }),
          new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:20,after:0},
            children:[new TextRun({text:`Fig. ${idx+1}`,italics:true,bold:true,size:14,font:'Calibri',color:'CCCCCC'})] })
        ];
      }
      cells.push(new TableCell({
        width:{size:Math.floor(TW/2),type:WidthType.DXA}, columnSpan:7,
        borders:allThin, verticalAlign:VerticalAlign.CENTER,
        margins:{top:40,bottom:40,left:60,right:60}, children
      }));
    }
    photoRows.push(new TableRow({ height:{value:2400}, children:cells }));
  }

  const mainTable = new Table({
    width:{size:TW,type:WidthType.DXA}, columnWidths:scaled,
    rows:[row_ig,row_sitio,row_dir,row_tk,row_tk2,row_sala,row_tec,
          row_rs,row_rs2,row_eq,row_eq_h,row_eq_d,
          row_med,row_med_h1,row_med_h2,row_med_d,row_obs,row_obs2]
  });

  const sectionChildren = [mainTable];
  if (photoRows.length > 0) {
    sectionChildren.push(
      new Paragraph({ pageBreakBefore:true, spacing:{before:0,after:0}, children:[] }),
      new Table({ width:{size:TW,type:WidthType.DXA}, columnWidths:scaled,
        rows:[row_foto_hdr,...photoRows] })
    );
  }

  return Packer.toBuffer(new Document({
    sections: [{ headers:{default:new Header({children:headerChildren})},
      properties:{page:{size:{width:12240,height:15840},
        margin:{top:1080,right:1701,bottom:1417,left:1701,header:284}}},
      children:sectionChildren }]
  }));
}

// ── Firestore helpers ─────────────────────────────────────
async function savePhotos(coleccion, docId, photos) {
  if (!photos || !photos.length) return;
  const batch = db.batch();
  photos.forEach((photo, idx) => {
    if (photo) {
      const ref = db.collection(coleccion).doc(docId).collection('photos').doc(String(idx));
      batch.set(ref, { data: photo, index: idx });
    }
  });
  await batch.commit();
}

async function loadPhotos(coleccion, docId) {
  const snap = await db.collection(coleccion).doc(docId).collection('photos').get();
  const photos = new Array(12).fill(null);
  snap.docs.forEach(d => { photos[d.data().index] = d.data().data; });
  return photos;
}

async function deletePhotos(coleccion, docId) {
  const snap = await db.collection(coleccion).doc(docId).collection('photos').get();
  if (snap.empty) return;
  const batch = db.batch();
  snap.docs.forEach(d => batch.delete(d.ref));
  await batch.commit();
}

// ── Routes ────────────────────────────────────────────────
app.get('/ping', (req,res) => res.json({ok:true}));

app.get('/version', (req,res) => {
  try {
    const mtime = fs.statSync(path.join(__dirname,'informe_clima_app.html')).mtimeMs;
    res.json({ v: mtime });
  } catch { res.json({ v: 0 }); }
});

app.post('/generar', async (req,res) => {
  try {
    const d = req.body;
    const buffer = await buildDocx(d);
    const sitePart = (d.nombreSitio||'Clima').replace(/[^a-zA-Z0-9]/g,'_').slice(0,25);
    const fname = `${d.codInforme||'Informe'}_${sitePart}.docx`;
    const docId = Date.now().toString();

    await db.collection('registros').doc(docId).set({
      fecha: d.fecha, fechaCreacion: new Date().toISOString(),
      codInforme: d.codInforme, nombreSitio: d.nombreSitio,
      codigoSitio: d.codigoSitio, direccion: d.direccion, sala: d.sala,
      tecnico: d.tecnico, supervisor: d.supervisor, numOT: d.numOT,
      ticketInc: d.ticketInc, ticketTE: d.ticketTE, ticketTI: d.ticketTI, ticketRED: d.ticketRED,
      resumen: d.resumen, observaciones: d.observaciones,
      eqSala: d.eqSala, eqNumero: d.eqNumero, eqTipo: d.eqTipo, eqMarca: d.eqMarca, eqModelo: d.eqModelo,
      m_cv: d.m_cv, m_ca: d.m_ca, m_ev: d.m_ev, m_ea: d.m_ea,
      m_condv: d.m_condv, m_conda: d.m_conda, m_tinj: d.m_tinj, m_tret: d.m_tret,
      photoCount: (d.photos||[]).filter(Boolean).length,
      photoDescs: d.photoDescs || [],
      filename: fname
    });
    await savePhotos('registros', docId, d.photos);

    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fname}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buffer);
  } catch(err) { console.error(err); res.status(500).json({error:err.message}); }
});

app.get('/registro', async (req,res) => {
  try {
    const q = (req.query.q||'').toLowerCase().trim();
    const snap = await db.collection('registros').orderBy('fechaCreacion','desc').get();
    let records = snap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    if (q) records = records.filter(r =>
      (r.nombreSitio||'').toLowerCase().includes(q) ||
      (r.codInforme||'').toLowerCase().includes(q) ||
      (r.tecnico||'').toLowerCase().includes(q) ||
      (r.numOT||'').toLowerCase().includes(q)
    );
    res.json(records);
  } catch(err) { res.status(500).json({error:err.message}); }
});

app.get('/descargar/:id', async (req,res) => {
  try {
    const doc = await db.collection('registros').doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({error:'No encontrado'});
    const data = doc.data();
    const photos = await loadPhotos('registros', req.params.id);
    const buffer = await buildDocx({ ...data, photos, photoDescs: data.photoDescs||[] });
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${data.filename}"`);
    res.send(buffer);
  } catch(err) { res.status(500).json({error:err.message}); }
});

app.post('/enviar/:id', async (req,res) => {
  try {
    const doc = await db.collection('registros').doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({error:'No encontrado'});
    const data = doc.data();
    const photos = await loadPhotos('registros', req.params.id);
    const buffer = await buildDocx({ ...data, photos, photoDescs: data.photoDescs||[] });
    const {to,smtpHost,smtpPort,smtpUser,smtpPass} = req.body;
    if (!to) return res.status(400).json({error:'Email requerido'});
    const t = nodemailer.createTransport({ host:smtpHost||'smtp.gmail.com', port:smtpPort||587, secure:false, auth:{user:smtpUser,pass:smtpPass} });
    await t.sendMail({ from:smtpUser, to,
      subject:`Informe - ${data.nombreSitio} - ${data.codInforme}`,
      text:`Adjunto informe.\nSitio: ${data.nombreSitio}\nFecha: ${data.fecha}\nTécnico: ${data.tecnico}`,
      attachments:[{filename:data.filename, content:buffer}]
    });
    res.json({ok:true});
  } catch(err){ res.status(500).json({error:err.message}); }
});

app.delete('/registro/:id', async (req,res) => {
  try {
    const doc = await db.collection('registros').doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({error:'No encontrado'});
    const data = doc.data();
    const photos = await loadPhotos('registros', req.params.id);
    await db.collection('papelera').doc(req.params.id).set({ ...data, fechaEliminado: new Date().toISOString() });
    await savePhotos('papelera', req.params.id, photos);
    await deletePhotos('registros', req.params.id);
    await db.collection('registros').doc(req.params.id).delete();
    res.json({ok:true});
  } catch(err) { res.status(500).json({error:err.message}); }
});

app.get('/papelera', async (req,res) => {
  try {
    const q = (req.query.q||'').toLowerCase().trim();
    const snap = await db.collection('papelera').orderBy('fechaEliminado','desc').get();
    let records = snap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    if (q) records = records.filter(r =>
      (r.nombreSitio||'').toLowerCase().includes(q) ||
      (r.codInforme||'').toLowerCase().includes(q) ||
      (r.tecnico||'').toLowerCase().includes(q)
    );
    res.json(records);
  } catch(err) { res.status(500).json({error:err.message}); }
});

app.post('/papelera/restaurar/:id', async (req,res) => {
  try {
    const doc = await db.collection('papelera').doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({error:'No encontrado'});
    const data = doc.data();
    const photos = await loadPhotos('papelera', req.params.id);
    const { fechaEliminado, ...clean } = data;
    await db.collection('registros').doc(req.params.id).set(clean);
    await savePhotos('registros', req.params.id, photos);
    await deletePhotos('papelera', req.params.id);
    await db.collection('papelera').doc(req.params.id).delete();
    res.json({ok:true});
  } catch(err) { res.status(500).json({error:err.message}); }
});

app.delete('/papelera/:id', async (req,res) => {
  try {
    await deletePhotos('papelera', req.params.id);
    await db.collection('papelera').doc(req.params.id).delete();
    res.json({ok:true});
  } catch(err) { res.status(500).json({error:err.message}); }
});

app.delete('/papelera', async (req,res) => {
  try {
    const snap = await db.collection('papelera').get();
    for (const doc of snap.docs) {
      await deletePhotos('papelera', doc.id);
      await doc.ref.delete();
    }
    res.json({ok:true});
  } catch(err) { res.status(500).json({error:err.message}); }
});

const PORT = process.env.PORT || 3000;
const os = require('os');
function getLocalIP() {
  for (const ifaces of Object.values(os.networkInterfaces()))
    for (const iface of ifaces)
      if (iface.family==='IPv4'&&!iface.internal) return iface.address;
  return 'localhost';
}
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Servidor corriendo en puerto ${PORT}`);
  console.log(`   Local: http://localhost:${PORT}`);
  if (PORT == 3000) console.log(`   Celular (misma red WiFi): http://${getLocalIP()}:${PORT}`);
});
