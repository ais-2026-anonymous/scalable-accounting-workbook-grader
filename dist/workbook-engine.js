/* OOXML patcher: preserves the user's native VBA project and form button.
   The spreadsheet export API does not export macro workbooks or their controls. */
export const SCHEMA='acc3000-instructor-v1';
const M='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P='http://schemas.openxmlformats.org/package/2006/relationships';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const X='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
export const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
export const COLORS=[{name:'Blue',hex:'DCEBFF'},{name:'Gold',hex:'FFF0BD'},{name:'Violet',hex:'E9DFFF'}];
const enc=new TextEncoder();
export const stable=v=>v===null||typeof v!=='object'?JSON.stringify(v):Array.isArray(v)?'['+v.map(stable).join(',')+']':'{'+Object.keys(v).sort().map(k=>JSON.stringify(k)+':'+stable(v[k])).join(',')+'}';
export async function sha(data){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',typeof data==='string'?enc.encode(data):data))].map(n=>n.toString(16).padStart(2,'0')).join('');}
// Normalize image responses consistently. Commitments remain only in private
// instructor records; new student workbooks contain no grading definition.
export const responseDigest=values=>sha(stable(values.map(v=>String(v??'').trim())));
// Grading definitions live in the instructor's browser/archive, never in new
// student workbooks. Older private archives can reconstruct their definition.
export function gradingDefinition(record){
 if(record.definition)return record.definition;
 const {schema,id,course,seed,labels,questions,factors,assignedId,checks,immutable}=record;
 return {schema,id,course,seed,labels,questions,factors,assignedId,checks:checks.map(({answer,target,imageFile,...publicFields})=>publicFields),immutable};
}
const xml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('Workbook contains invalid XML.');return d;};
const str=d=>new XMLSerializer().serializeToString(d);
const all=(n,tag,ns=M)=>[...n.getElementsByTagNameNS(ns,tag)];
const first=(n,tag,ns=M)=>all(n,tag,ns)[0];
function resolve(base,target){const p=target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target;const a=[];for(const x of p.split('/')){if(x==='..')a.pop();else if(x!=='.'&&x)a.push(x);}return a.join('/');}
function relPath(p){return p.slice(0,p.lastIndexOf('/')+1)+'_rels/'+p.slice(p.lastIndexOf('/')+1)+'.rels';}
async function text(z,p){const f=z.file(p);if(!f)throw Error('Required workbook part is missing: '+p);return f.async('string');}
export async function openBook(bytes,name='workbook.xlsm'){
 if(!/\.xlsm$/i.test(name))throw Error('Use a genuine .xlsm macro workbook for this instructor prototype.');
 if(bytes.byteLength>20*1024*1024)throw Error('Workbook exceeds 20 MB.');
 let z;try{z=await globalThis.JSZip.loadAsync(bytes);}catch{throw Error('The file is not a valid Excel workbook. Upload the original .xlsm file saved from Excel.');}
 const entries=Object.values(z.files);if(entries.length>4000||entries.reduce((s,f)=>s+(f._data?.uncompressedSize??0),0)>100*1024*1024)throw Error('Workbook exceeds extraction limits.');
 if(entries.some(f=>f.name.includes('externalLinks/')||f.unsafeOriginalName?.split('/').includes('..')))throw Error('External links or unsafe archive paths are not accepted.');
 const types=await text(z,'[Content_Types].xml');
 if(!types.includes('application/vnd.ms-excel.sheet.macroEnabled.main+xml')||!z.file('xl/vbaProject.bin'))throw Error('The file is not a genuine macro-enabled Excel workbook.');
 const book=xml(await text(z,'xl/workbook.xml')),rels=xml(await text(z,'xl/_rels/workbook.xml.rels'));
 const targets=new Map(all(rels,'Relationship',P).map(r=>[r.getAttribute('Id'),resolve('xl/workbook.xml',r.getAttribute('Target'))]));
 const sheets=new Map(all(book,'sheet').map(s=>[s.getAttribute('name'),{node:s,path:targets.get(s.getAttributeNS(R,'id'))}]));
 const strings=z.file('xl/sharedStrings.xml')?all(xml(await text(z,'xl/sharedStrings.xml')),'si').map(si=>all(si,'t').map(t=>t.textContent).join('')):[];
 const docs=new Map();for(const [name,s] of sheets)docs.set(name,xml(await text(z,s.path)));
 const styles=xml(await text(z,'xl/styles.xml'));
 const date1904=first(book,'workbookPr')?.getAttribute('date1904');
 if(date1904==='1'||date1904==='true')throw Error('Use the supplied 1900-date-system case workbook.');
 return {z,book,rels,sheets,docs,strings,styles};
}
function cells(doc,strings){return new Map(all(doc,'c').map(c=>{
 const t=c.getAttribute('t'),v=first(c,'v')?.textContent;
 let value=t==='s'?strings[Number(v)]:t==='inlineStr'?all(c,'t').map(t=>t.textContent).join(''):t==='str'||t==='e'?v:v==null?null:Number(v);
 return[c.getAttribute('r'),{value,formula:first(c,'f')?.textContent,node:c,type:t}];
}));}
function dateKey(v,monthly=false){
 if(typeof v==='number')return new Date(Math.round((v-25569)*86400000)).toISOString().slice(0,monthly?7:10);
 if(/^\d{4}-\d{2}/.test(String(v)))return String(v).slice(0,monthly?7:10);
 const m=String(v).match(/^([A-Za-z]+) (\d{4})$/),i=m?MONTHS.indexOf(m[1]):-1;
 if(i<0)throw Error('Unrecognized source date.');return m[2]+'-'+String(i+1).padStart(2,'0');
}
function sourceRows(b,name,end,monthly=false){
 if(!b.docs.has(name))throw Error('Missing case source sheet: '+name);
 const c=cells(b.docs.get(name),b.strings),rows=[];
 for(let r=6;r<=end;r++){
  const row=[dateKey(c.get('A'+r)?.value,monthly)];
  for(const col of ['B','C','D','E']){
   const item=c.get(col+r),v=item?.value;
   if(item?.formula)throw Error('Source data still contain formulas. Run the supplied Generate Data button and save.');
   if(v!==null&&v!==undefined&&(typeof v!=='number'||!Number.isFinite(v)||v<0))throw Error('Invalid numeric source data.');
   row.push(v??null);
  }rows.push(row);
 }
 return rows;
}
async function pictureChecks(b){
 const solution=b.docs.get('Solution'),ref=first(solution,'drawing');if(!ref)throw Error('The image instruction drawing is missing.');
 const rels=xml(await text(b.z,relPath(b.sheets.get('Solution').path)));
 const r=all(rels,'Relationship',P).find(r=>r.getAttribute('Id')===ref.getAttributeNS(R,'id'));
 if(!r)throw Error('Image drawing relationship is missing.');
 const path=resolve(b.sheets.get('Solution').path,r.getAttribute('Target')),d=xml(await text(b.z,path)),dr=xml(await text(b.z,relPath(path)));
 const targets=new Map(all(dr,'Relationship',P).map(r=>[r.getAttribute('Id'),resolve(path,r.getAttribute('Target'))]));
 const checks=[];for(const pic of all(d,'pic',X)){
  const anchor=pic.parentNode,from=first(anchor,'from',X),blip=first(pic,'blip',A),prop=first(pic,'cNvPr',X);
  if(!from||!blip||prop?.getAttribute('hidden')==='1')throw Error('The instruction image is missing, hidden or relocated.');
  const target=targets.get(blip.getAttributeNS(R,'embed'));
  checks.push({row:Number(first(from,'row',X).textContent)+1,col:Number(first(from,'col',X).textContent),imageSha:await sha(await b.z.file(target).async('uint8array'))});
 }return checks.sort((a,b)=>a.row-b.row);
}
export async function verifySubmission(bytes,name,recorded){
 const receipt=recorded,expected=receipt?.checksum;
 if(receipt?.schema!==SCHEMA||!/^[a-f0-9]{64}$/.test(expected??''))throw Error('The instructor’s private grading record is required.');
 const b=await openBook(bytes,name),meta=gradingDefinition(receipt);
 if(meta.schema!==SCHEMA||meta.questions?.length!==10||!Array.isArray(meta.checks))throw Error('This workbook has an unsupported assignment definition.');
 if(receipt&&meta.id!==receipt.id)throw Error('Assignment version mismatch: this submission belongs to '+meta.id+'.');
 if(b.sheets.has('Reference')||b.sheets.has('Title Page'))throw Error('Run Generate Data once in desktop Excel and save the resulting workbook before grading.');
 for(const n of ['Solution','Sales Locked','Warranties & Returns Locked'])if(!first(b.docs.get(n)??xml('<missing/>'),'sheetProtection'))throw Error('Required sheet protection is missing: '+n);
 const sol=cells(b.docs.get('Solution'),b.strings);
 if(sol.get('B5')?.value!==expected)throw Error('The workbook checksum does not match the instructor’s recorded checksum.');
 if(meta.immutable.some(([a])=>sol.get(a)?.formula))throw Error('Assignment labels must remain fixed text.');
 for(const name of ['Sales Locked','Warranties & Returns Locked']){
  const source=cells(b.docs.get(name),b.strings);
  if(source.get('B4')?.value!==meta.labels.A.replace(/^Product /,'')||source.get('D4')?.value!==meta.labels.B.replace(/^Product /,''))throw Error('Checksum mismatch: protected product identifiers have changed.');
 }
 const position=a=>{const m=a.match(/^([A-Z]+)(\d+)$/);return [[...m[1]].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0),Number(m[2])];};
 const answerCells=[...meta.questions.map(q=>q.cell),...meta.checks.flatMap(c=>c.boxes.map(b=>b.cell))];
 for(const merge of all(b.docs.get('Solution'),'mergeCell')){
  const [start,end]=merge.getAttribute('ref').split(':'),[c1,r1]=position(start),[c2,r2]=position(end??start);
  if(answerCells.some(a=>{const[c,r]=position(a);return c>=c1&&c<=c2&&r>=r1&&r<=r2;}))throw Error('An answer box has been merged or relocated.');
 }
 const styles=all(b.styles,'cellXfs')[0],fills=all(b.styles,'fills')[0];
 for(const check of meta.checks)for(const box of check.boxes){
  const si=Number(sol.get(box.cell)?.node.getAttribute('s')??0),xf=styles.children[si],fill=fills.children[Number(xf?.getAttribute('fillId')??0)];
  if(first(fill,'fgColor')?.getAttribute('rgb')!=='FF'+box.hex)throw Error('The colored answer-box layout has changed.');
 }
 const immutable=meta.immutable.map(([a])=>[a,sol.get(a)?.value??null]);
 if(stable(immutable)!==stable(meta.immutable))throw Error('Checksum mismatch: assignment labels or question positions have changed.');
 // Earlier versions issued a fixed data snapshot. New versions authenticate
 // the assignment definition while Python validates each student's own data.
 let sourceDigest=null;
 if(meta.dataMode!=='student-generated'){
  const data={sales:sourceRows(b,'Sales Locked',3658),claims:sourceRows(b,'Warranties & Returns Locked',125,true)};
  sourceDigest=await sha(stable(data));
 }
 const images=await pictureChecks(b);
 if(images.length!==meta.checks.length||images.some((v,i)=>v.row!==meta.checks[i].row||v.col!==0||v.imageSha!==meta.checks[i].imageSha))throw Error('Checksum mismatch: instruction images are missing, changed or relocated.');
 const checksum=await sha(stable({meta,sourceDigest,images:images.map(i=>i.imageSha)}));
 if(checksum!==expected||(receipt&&sourceDigest!==receipt.sourceDigest))throw Error('Checksum mismatch: the assignment differs from the issued version.');
 const visual=await Promise.all((receipt?.checks??meta.checks).map(async check=>{
  const actual=check.boxes.map(box=>sol.get(box.cell)?.value??null);
  const passed=check.responseDigest?await responseDigest(actual)===check.responseDigest:receipt?String(actual[check.target]??'').trim()===check.answer&&actual.every((v,i)=>i===check.target||v===null||String(v).trim()===''):null;
  return {id:check.id,passed,expectedCell:check.boxes.map(b=>b.cell).join(' / '),actual,message:passed===null?'Older workbook: image grading needs a newly generated version. Technical grading is available.':passed?'Image instruction completed.':'The value must be in the indicated box and the other two boxes must be empty.'};
 }));
 return {valid:true,checksum,meta,visual,identity:{last:sol.get('B1')?.value,first:sol.get('B2')?.value,id:sol.get('B3')?.value},questions:meta.questions};
}
