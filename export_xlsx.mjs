import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
const [input, output] = process.argv.slice(2);
if (!input || !output) throw Error('Input dan output diperlukan.');
const records = JSON.parse(await fs.readFile(input,'utf8'));
if (!records.length) throw Error('Tidak ada hasil.');
const columns = Object.keys(records[0]);
if (columns.at(-1)!=='kategori_pelaporan') throw Error('Kolom terakhir harus kategori_pelaporan.');
const wb=Workbook.create();
const sh=wb.worksheets.add('Laporan WBS');
const rows=records.map(r=>columns.map(c=>{
 const v=r[c]??null;
 return typeof v==='string' && v.startsWith('=') ? "'"+v : v;
}));
sh.getRangeByIndexes(0,0,rows.length+1,columns.length).values=[columns,...rows];
const all=sh.getRangeByIndexes(0,0,rows.length+1,columns.length);
all.format.font={name:'Calibri',size:11};
all.format.verticalAlignment='top';
sh.getRangeByIndexes(0,0,1,columns.length).format={fill:'#173F45',font:{bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:36};
sh.getRangeByIndexes(1,0,rows.length,columns.length).format.rowHeight=76;
for(let i=0;i<columns.length;i++){
 const col=sh.getRangeByIndexes(0,i,rows.length+1,1);
 col.format.columnWidth=columns[i]==='Detail Laporan'?78:columns[i]==='kategori_pelaporan'?52:24;
 col.format.wrapText=true;
}
sh.freezePanes.freezeRows(1);
sh.freezePanes.freezeColumns(2);
sh.tables.add(all,true,'LaporanKategori');
wb.recalculate();
if(sh.getRangeByIndexes(0,0,1,columns.length).values[0].at(-1)!=='kategori_pelaporan')throw Error('Header mismatch');
const cat=columns.length-1;
const preview=await wb.render({sheetName:'Laporan WBS',range:`${String.fromCharCode(65+cat)}1:${String.fromCharCode(65+cat)}${Math.min(rows.length+1,7)}`,scale:1.3,format:'png'});
await fs.writeFile(output+'.preview.png',new Uint8Array(await preview.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(JSON.stringify({output,rows:records.length,columns:columns.length}));