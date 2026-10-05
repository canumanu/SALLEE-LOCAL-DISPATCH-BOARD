/* ---- Local chart parser (Excel -> drivers + jobs). Pure function: takes a SheetJS worksheet. ---- */
function parseLocalChart(ws, opts){
  opts = opts || {};
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:true});
  const cell = (r,c)=>{ const row=rows[r-1]; return row && row[c-1]!==undefined ? row[c-1] : null; };
  const isStr = v=>typeof v==='string' && v.trim()!=='';
  const isTime = v=>typeof v==='number' && v>0 && v<1;
  const result = {drivers:[], jobs:[], warnings:[], firstDate:null, lastDate:null, days:0, statusCount:0, noteCount:0, skipped:{}};

  /* --- driver column pairs (row 1 names, row 3 type) --- */
  const TYPE_MAP = (t)=>{
    const s=String(t||'').toUpperCase().replace(/\s+/g,' ').trim();
    if(s==='LOCAL / OTR' || s==='LOCAL/OTR') return 'LOCAL_OTR';
    if(s.startsWith('AGENT')) return 'AGENT';
    if(s==='PT') return 'PT';
    return 'LOCAL'; // LOCAL, and anything else on a local chart
  };
  let emptyRun=0;
  for(let c=2;c<=80;c+=2){
    const name=cell(1,c);
    if(isStr(name)){
      emptyRun=0;
      const truckCell=cell(1,c+1);
      let truck='',trailer='';
      if(typeof truckCell==='number'){ truck=String(truckCell); }
      else if(isStr(truckCell)){ const parts=truckCell.split('/'); truck=parts[0].trim(); trailer=(parts[1]||'').trim(); }
      result.drivers.push({col:c, name1:name.trim(), name2:'', truck, trailer, location:TYPE_MAP(cell(3,c)), equipment:String(cell(2,c+1)||'').trim()});
    } else {
      emptyRun++;
      if(emptyRun>=3 && result.drivers.length) break;
    }
  }
  if(!result.drivers.length){ result.warnings.push('No driver columns found in row 1.'); return result; }

  /* --- day blocks: weekday labels (MON/TUE..) in column A, date in the row below.
         Rows get inserted on busy days, so spacing between labels varies (7-31 rows).
         A day's block starts 7 rows above its label and runs until the next day's block starts. --- */
  const DOW=['SUN','MON','TUE','WED','THU','FRI','SAT'];
  const labelRows=[];
  // a "label" = a short word in column A (MON, TUES, THURS... tolerates typos like "TUH"; the date may sit above or below it)
  for(let r=4;r<=rows.length;r++){ const v=cell(r,1); if(isStr(v) && /^[A-Za-z]{3,6}$/.test(v.trim())) labelRows.push(r); }
  if(!labelRows.length){ result.warnings.push('Could not find any day labels (MON/TUE/...) in column A.'); return result; }
  const firstSerial=typeof cell(labelRows[0]+1,1)==='number' ? cell(labelRows[0]+1,1) : cell(labelRows[0]-1,1);
  if(typeof firstSerial!=='number'){ result.warnings.push('Could not read the first date under the first day label.'); return result; }
  const p=XLSX.SSF.parse_date_code(firstSerial);
  result.parsedStart=new Date(Date.UTC(p.y,p.m-1,p.d)).toISOString().slice(0,10);
  let baseUTC=Date.UTC(p.y,p.m-1,p.d);
  if(opts.startDate){ const [y,m,d]=opts.startDate.split('-').map(Number); baseUTC=Date.UTC(y,m-1,d); }
  const dateStrAt=i=>new Date(baseUTC+i*86400000).toISOString().slice(0,10);
  const dowAt=i=>new Date(baseUTC+i*86400000).getUTCDay();
  result.firstLabel=cell(labelRows[0],1).trim().toUpperCase().slice(0,3);
  if(DOW[dowAt(0)]!==result.firstLabel){ result.warnings.push(`The first day in this chart is labelled ${result.firstLabel}, but ${dateStrAt(0)} is a ${DOW[dowAt(0)]}. Pick a first date that is a ${result.firstLabel}.`); return result; }

  // Consecutive days = labels that follow the weekday sequence. A stray typo is warned about and tolerated;
  // 3 wrong labels in a row means we've run into a stale/leftover copy, so stop there.
  let days=0,bad=0;
  for(let i=0;i<labelRows.length;i++){
    const lab=cell(labelRows[i],1).trim().toUpperCase();
    if(lab.slice(0,3)===DOW[dowAt(i)]){ bad=0; days=i+1; }
    else{
      bad++;
      if(bad>=3){ days=i-2; break; }
      days=i+1;
      result.warnings.push(`Row ${labelRows[i]}: label "${lab}" where ${DOW[dowAt(i)]} was expected (counted as ${DOW[dowAt(i)]} ${dateStrAt(i)}).`);
    }
  }
  result.days=days;
  result.firstDate=dateStrAt(0); result.lastDate=dateStrAt(days-1);
  const blockStart=i=>Math.max(4,labelRows[i]-7);
  const blockEnd=i=>(i+1<days)?labelRows[i+1]-8:labelRows[i]+7;

  /* --- time helpers --- */
  const hhmm=m=>String(Math.floor(m/60)%24).padStart(2,'0')+':'+String(m%60).padStart(2,'0');
  const fromFraction=v=>hhmm(Math.round(v*1440));
  const pretty=t=>{ const [h,m]=t.split(':').map(Number); const ap=h>=12?'PM':'AM'; return `${((h+11)%12)+1}:${String(m).padStart(2,'0')} ${ap}`; };
  function parseTimeCell(v){
    if(isTime(v)) return {time:fromFraction(v)};
    if(isStr(v)){
      const s=v.trim().toUpperCase().replace(/O/g,'0');
      const m=s.match(/^(\d{1,2})(?::(\d{2}))?:?\s*(AM|PM)?$/);
      if(m){ let h=+m[1], mi=+(m[2]||0); if(m[3]==='PM'&&h<12)h+=12; if(m[3]==='AM'&&h===12)h=0; if(h<24&&mi<60) return {time:hhmm(h*60+mi)}; }
      return {raw:v.trim()};
    }
    return {};
  }
  const txt=v=>isStr(v)?v.trim().replace(/\s+/g,' '):'';
  const STATUS=new Set(['UNPAID','OFF','VACATION','PERSONAL','SICK','HOME','NO SHOW / NO CALL']);
  const HEADER_WORDS=new Set(['KY','KY /LA','LOCAL','LOCAL / OTR','LOCAL/OTR','AGENT/DRIVER','PT','GN','6H','3H','BV','LA']);
  // second line of a job group: DROP, RACE & RETURN, BREEDING TRIP, NIGHT BREEDING TRIP, SHUTTLE, WAIT & RETURN, SWAP ...
  const JOBLINE=/DROP|RACE|SWAP|BREEDING|SHUTTLE|WAIT|RETURN|TRIP|PICK ?UP|DELIVER/i;
  const KIND=k=>{
    k=k.toUpperCase();
    if(k.includes('NIGHT') && k.includes('BREEDING')) return 'night_breeding';
    if(k.includes('BREEDING')) return 'breeding';
    if(k.includes('SHUTTLE')) return 'shuttle';
    if(k.includes('WAIT')) return 'wait_return';
    if(k.includes('RACE')) return 'race_and_return';
    if(k.includes('SWAP')) return 'swap';
    if(/^DROP\b/.test(k)) return 'drop';
    return 'other';
  };

  /* --- walk each driver, each day block (two passes: find job groups, then deal with leftovers) --- */
  const prettyT=v=>pretty(fromFraction(v));
  const stripTag=n=>n.toUpperCase().replace(/#\d+/g,'').replace(/[^A-Z0-9]/g,'');
  const driverKeys=new Set(result.drivers.map(d=>stripTag(d.name1)));
  const looksLikeTruck=v=>/^\d{3}\s*\/\s*\d+[A-Z+]*(\s*\(.*\))?$/i.test(v) || /^P\d+\s*\/\s*\d+/i.test(v);
  result.drivers.forEach(drv=>{
    const c=drv.col;
    for(let i=0;i<days;i++){
      const date=dateStrAt(i);
      const start=blockStart(i), end=blockEnd(i);
      let order=0;
      // pass 1: job groups = customer row, type line, origin/destination row
      const groups=[];
      for(let r=start;r<=end;){
        const a=cell(r,c);
        if(isStr(a)){
          const label=txt(a);
          if(stripTag(label)===stripTag(drv.name1)){ r+=3; continue; } // weekly header repeat (name / type / LOCAL rows)
          const b=cell(r+1,c);
          if(!STATUS.has(label.toUpperCase()) && isStr(b) && b.trim().length<=30 && JOBLINE.test(b.trim())){ groups.push(r); r+=3; continue; }
        }
        r++;
      }
      const inGroup=new Set(); groups.forEach(g=>{inGroup.add(g);inGroup.add(g+1);inGroup.add(g+2);});
      const tagRows=new Set(); // preamble rows consumed by a job below them
      let prevEnd=start-1;
      groups.forEach(r=>{
        const label=txt(cell(r,c)), b=cell(r+1,c);
        const when=parseTimeCell(cell(r,c+1));
        const job={date,driverCol:c,order:0,type:'normal',customer:label,
          origin:txt(cell(r+2,c)),destination:txt(cell(r+2,c+1)),time:when.time||'',timeIn:'',timeOut:'',notes:'',loadType:KIND(b.trim())};
        const bits=[];
        if(job.loadType==='other' || /CANCEL/i.test(b)) bits.push(txt(b));
        if(when.raw) bits.push('Time: '+when.raw);
        const lineNote=txt(cell(r+1,c+1)); if(lineNote) bits.push(lineNote);
        // preamble just above the job: row r-1 left = IN time (told to report), row r-2 = OUT time or "WD" (when done)
        const pre=[];
        const isWD=v=>isStr(v) && /^\s*W\.?D\.?\s*$/i.test(v);
        for(let rr=Math.max(prevEnd+1,start,r-3); rr<=r-1; rr++){
          const L=cell(rr,c), R=cell(rr,c+1);
          if(rr===r-1){
            if(isTime(L)){ job.timeIn=fromFraction(L); tagRows.add(rr); }
            if(isTime(R)){ pre.push('Also '+prettyT(R)); tagRows.add(rr); }
          } else if(rr===r-2){
            let o=null;
            if(isWD(L)||isTime(L)){ o=L; if(isTime(R)) pre.push('Also '+prettyT(R)); }
            else if(isWD(R)||isTime(R)){ o=R; }
            if(o!==null){ job.timeOut = isWD(o) ? 'WD' : fromFraction(o); tagRows.add(rr); }
            if(isStr(R) && !isWD(R)){ const t=txt(R); pre.push(looksLikeTruck(t)?'Truck '+t:t); tagRows.add(rr); }
          } else {
            if(isStr(R)){ const t=txt(R); if(!isWD(R)){ pre.push(looksLikeTruck(t)?'Truck '+t:t); tagRows.add(rr); } }
            if(isTime(R)){ pre.push(prettyT(R)); tagRows.add(rr); }
          }
        }
        // out times are written without AM/PM (3:30 means 3:30 PM): move them to the afternoon when needed
        if(job.timeOut && job.timeOut!=='WD'){
          let [oh,om]=job.timeOut.split(':').map(Number);
          const inMin=job.timeIn ? (+job.timeIn.slice(0,2))*60+(+job.timeIn.slice(3)) : null;
          const outMin=oh*60+om;
          if(oh<12 && (inMin===null ? oh<=7 : outMin<=inMin)) job.timeOut=hhmm(outMin+720);
        }
        if(pre.length) bits.push(pre.join(' · '));
        job.notes=bits.join(' · ');
        job._row=r;
        prevEnd=r+2;
        job.order=0;
        result.jobs.push(job);
      });
      // pass 2: everything else in this block for this driver
      for(let r=start;r<=end;r++){
        if(inGroup.has(r)) continue;
        const a=cell(r,c); if(!isStr(a)) continue;
        const label=txt(a), U=label.toUpperCase();
        if(tagRows.has(r) && label.length<=4) continue;               // "WD"-style tag already attached to a job
        if(stripTag(label)===stripTag(drv.name1)) { r+=2; continue; } // header repeat
        if(driverKeys.has(stripTag(label)) || HEADER_WORDS.has(U) || looksLikeTruck(label)) continue;
        if(U==='CALLED' && isStr(cell(r+1,c)) && driverKeys.has(stripTag(cell(r+1,c)))) continue;
        if(STATUS.has(U)){ result.jobs.push({date,driverCol:c,order:0,type:'vacation',notes:label}); result.statusCount++; continue; }
        if(U==='WD' || U==='CALLED'){ result.skipped[U]=(result.skipped[U]||0)+1; continue; } // layout markers, counted but not imported as cards
        result.jobs.push({date,driverCol:c,order:0,type:'note',notes:label}); result.noteCount++;
      }
    }
  });
  // stable order within each driver/day: by sheet position for jobs, then statuses/notes
  const cnt={};
  result.jobs.forEach(j=>{ const k=j.date+'|'+j.driverCol; cnt[k]=(cnt[k]||0)+1; j.order=cnt[k]; delete j._row; });
  return result;
}
if(typeof module!=='undefined') module.exports={parseLocalChart};
