import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { ArrowLeft, ChevronDown, ChevronRight, ChevronUp, House, Search, Settings, Upload, Volume2, X } from "lucide-react";
import hornet from "../design/silksong/characters/03_hornet.png";
import sherma from "../design/silksong/characters/01_sherma.png";
import lace from "../design/silksong/characters/04_lace.png";
import phantom from "../design/silksong/characters/07_phantom.png";
import garmond from "../design/silksong/characters/06_garmond_zaza.png";
import pilgrim from "../design/silksong/characters/02_pilgrim_shield.png";
import cloak from "../design/silksong/characters/08_tall_cloak.png";
import group from "../design/silksong/characters/09_pilgrim_group.png";
import nuu from "../design/silksong/characters/11_nuu.png";
import grindle from "../design/silksong/characters/12_grindle.png";

type Screen = "home"|"import"|"character"|"quiz"|"result"|"journal";
type Mode = "Dictation"|"C → E"|"E → C";
type Allocation = "today"|"split";
type Character = {name:string; src:string};
type Word = {en:string; zh:string; answer:string; status:"correct"|"wrong"|"skip"};
type ImportedWord = {en:string; zh:string; date:string|null};
type Attempt = ImportedWord & {answer:string; status:"correct"|"wrong"|"skip"; round:number};
type Session = {id:number; mode:Mode; character:Character; attempts:Attempt[]; date:string};

const characters: Character[] = [
  {name:"Sherma",src:sherma},{name:"Hornet",src:hornet},{name:"Lace",src:lace},
  {name:"Phantom",src:phantom},{name:"Garmond & Zaza",src:garmond},{name:"Pilgrim",src:pilgrim},
  {name:"Tall Cloak",src:cloak},{name:"Pilgrim Group",src:group},{name:"Nuu",src:nuu},{name:"Grindle",src:grindle}
];
const results: Word[] = [
  {en:"apple",zh:"苹果",answer:"apple",status:"correct"},{en:"orange",zh:"橙子；橙色",answer:"orange",status:"correct"},
  {en:"beautiful",zh:"美丽的",answer:"beatiful",status:"wrong"},{en:"tomorrow",zh:"明天",answer:"tomorrow",status:"correct"},
  {en:"necessary",zh:"必要的",answer:"neccessary",status:"wrong"},{en:"family",zh:"家庭",answer:"family",status:"correct"},
  {en:"knowledge",zh:"知识",answer:"—",status:"skip"},{en:"library",zh:"图书馆",answer:"library",status:"correct"}
];
const months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];

function readStored<T>(key:string,fallback:T):T{
  try{const value=localStorage.getItem(key);return value?JSON.parse(value) as T:fallback}catch{return fallback}
}

export function App(){
  const [screen,setScreen]=useState<Screen>("home");
  const [mode,setMode]=useState<Mode>("Dictation");
  const [allocation,setAllocation]=useState<Allocation>("split");
  const [selected,setSelected]=useState(1);
  const [answer,setAnswer]=useState("");
  const [question,setQuestion]=useState(1);
  const [reviewOnly,setReviewOnly]=useState(false);
  const [pendingWords,setPendingWords]=useState<ImportedWord[]>([]);
  const [importedWords,setImportedWords]=useState<ImportedWord[]>(()=>readStored("englishc.words",[]));
  const [fileName,setFileName]=useState("");
  const [importError,setImportError]=useState("");
  const [activeWords,setActiveWords]=useState<ImportedWord[]>([]);
  const [attempts,setAttempts]=useState<Attempt[]>([]);
  const [round,setRound]=useState(1);
  const [sessions,setSessions]=useState<Session[]>(()=>readStored<Session[]>("englishc.sessions",[]).map(session=>({...session,character:characters.find(character=>character.name===session.character?.name)??characters[0]})));
  useEffect(()=>{localStorage.setItem("englishc.words",JSON.stringify(importedWords))},[importedWords]);
  useEffect(()=>{localStorage.setItem("englishc.sessions",JSON.stringify(sessions))},[sessions]);
  const chosen=characters[selected];
  const go=(s:Screen)=>setScreen(s);
  return <main className="stage"><section className="phone">
    {screen==="home"&&<Home mode={mode} setMode={setMode} go={go} importedWords={importedWords} sessions={sessions} selectDay={(day)=>{const date=`2026-08-${String(day).padStart(2,"0")}`;setActiveWords(importedWords.filter(w=>w.date===date));setQuestion(1);setRound(1);setAnswer("");setAttempts([]);setReviewOnly(false);go("character")}} onFile={(words,name)=>{setPendingWords(words);setFileName(name);setImportError("");go("import")}} importError={importError} setImportError={setImportError}/>} 
    {screen==="import"&&<Import allocation={allocation} setAllocation={setAllocation} go={go} words={pendingWords} fileName={fileName} confirm={(words)=>{setImportedWords(words);setImportError("");go("home")}}/>} 
    {screen==="character"&&<Characters selected={selected} setSelected={setSelected} go={go}/>} 
    {screen==="quiz"&&<Quiz mode={mode} chosen={chosen} answer={answer} setAnswer={setAnswer} question={question} setQuestion={setQuestion} go={go} reviewOnly={reviewOnly} round={round} words={activeWords} attempts={attempts} setAttempts={setAttempts}/>} 
    {screen==="result"&&<Result chosen={chosen} mode={mode} go={go} round={round} attempts={attempts} done={()=>{const date=attempts[0]?.date??"2026-08-12";setSessions(current=>[{id:Date.now(),mode,character:chosen,attempts:[...attempts],date},...current]);go("journal")}} retry={()=>{const missed=attempts.filter(a=>a.round===1&&a.status!=="correct").map(({en,zh,date})=>({en,zh,date}));setActiveWords(missed);setRound(r=>r+1);setReviewOnly(true);setQuestion(1);setAnswer("");go("quiz")}}/>} 
    {screen==="journal"&&<Journal go={go} sessions={sessions}/>} 
  </section></main>
}

function IconButton({children,onClick,label}:{children:React.ReactNode,onClick:()=>void,label:string}){return <button className="icon" onClick={onClick} aria-label={label}>{children}</button>}

function Home({mode,setMode,go,importedWords,sessions,selectDay,onFile,importError,setImportError}:{mode:Mode;setMode:(m:Mode)=>void;go:(s:Screen)=>void;importedWords:ImportedWord[];sessions:Session[];selectDay:(d:number)=>void;onFile:(w:ImportedWord[],name:string)=>void;importError:string;setImportError:(s:string)=>void}){
  const swipeStart=useRef<{x:number;y:number}|null>(null);
  const [menuOpen,setMenuOpen]=useState(false);
  const [monthPicker,setMonthPicker]=useState(false);
  const [year,setYear]=useState(2026);
  const [month,setMonth]=useState(7);
  const currentDate=new Date();
  const isCurrentMonth=year===currentDate.getFullYear()&&month===currentDate.getMonth();
  const today=isCurrentMonth?currentDate.getDate():null;
  const dayCount=new Date(year,month+1,0).getDate();
  const leadingBlanks=new Date(year,month,1).getDay();
  const calendarCells=[...Array.from({length:leadingBlanks},()=>null),...Array.from({length:dayCount},(_,i)=>i+1)];
  const availableDays=Array.from(new Set(importedWords.filter(w=>w.date?.startsWith(`${year}-${String(month+1).padStart(2,"0")}`)).map(w=>Number(w.date!.slice(-2))))).sort((a,b)=>a-b);
  const [selectedDay,setSelectedDay]=useState<number|null>(today&&availableDays.includes(today)?today:(availableDays[0]??null));
  const fileRef=useRef<HTMLInputElement>(null);
  const readFile=async(file?:File)=>{
    if(!file)return;
    try{
      // Keep Excel dates as serial values. Converting them to JavaScript Date
      // objects can shift the calendar day when the workbook timezone differs.
      const book=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:false});
      const sheet=book.Sheets[book.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json<Record<string,unknown>>(sheet,{defval:null,raw:true});
      const normalizeDate=(value:unknown):string|null=>{
        if(value===null||value===undefined||value==="")return null;
        if(value instanceof Date&&!isNaN(value.getTime()))return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,"0")}-${String(value.getDate()).padStart(2,"0")}`;
        if(typeof value==="number"){const p=XLSX.SSF.parse_date_code(value);if(p)return `${p.y}-${String(p.m).padStart(2,"0")}-${String(p.d).padStart(2,"0")}`}
        const text=String(value).trim().replaceAll("/","-");
        return /^\d{4}-\d{1,2}-\d{1,2}$/.test(text)?text.split("-").map((v,i)=>i?String(Number(v)).padStart(2,"0"):v).join("-"):"INVALID";
      };
      const parsed=rows.map((r,i)=>({en:String(r["English Word *"]??r["English Word"]??"").trim(),zh:String(r["Chinese Meaning *"]??r["Chinese Meaning"]??"").trim(),date:normalizeDate(r["Dictation Date (Optional)"]??r["Dictation Date"]),row:i+2}));
      const invalid=parsed.filter(r=>!r.en||!r.zh||r.date==="INVALID");
      if(!parsed.length)throw new Error("No words were found. Please use the provided template.");
      if(invalid.length)throw new Error(`Please review row${invalid.length>1?"s":""} ${invalid.slice(0,5).map(r=>r.row).join(", ")}. English, Chinese, or date is invalid.`);
      onFile(parsed.map(({en,zh,date})=>({en,zh,date})),file.name);
    }catch(e){setImportError(e instanceof Error?e.message:"The file could not be read.")}
    if(fileRef.current)fileRef.current.value="";
  };
  const importedDays=new Set(availableDays);
  const completedByDay=new Map(sessions.filter(s=>s.date.startsWith(`${year}-${String(month+1).padStart(2,"0")}`)).map(s=>[Number(s.date.slice(-2)),s.character]));
  return <div className="page home" onPointerDown={e=>{swipeStart.current={x:e.clientX,y:e.clientY}}} onPointerCancel={()=>{swipeStart.current=null}} onPointerUp={e=>{const start=swipeStart.current;swipeStart.current=null;if(!start)return;const dx=e.clientX-start.x,dy=e.clientY-start.y;if(dx>65&&Math.abs(dx)>Math.abs(dy)*1.25)go("journal")}}><div className="top"><button className="home-menu-toggle" onClick={()=>setMenuOpen(!menuOpen)} aria-expanded={menuOpen} aria-label={menuOpen?"Collapse test modes":"Expand test modes"}><House/>{menuOpen?<ChevronUp/>:<ChevronDown/>}</button><IconButton onClick={()=>{}} label="Settings"><Settings/></IconButton></div>
    {menuOpen&&<div className="mode">{(["Dictation","C → E","E → C"] as Mode[]).map(m=><button className={mode===m?"on":""} onClick={()=>{setMode(m);setMenuOpen(false)}} key={m}>{m}</button>)}</div>}
    <button className="month" onClick={()=>setMonthPicker(true)}><small>{year}</small><h1>{months[month]==="SEP"?"SEPTEMBER":months[month]==="OCT"?"OCTOBER":months[month]==="NOV"?"NOVEMBER":months[month]==="DEC"?"DECEMBER":months[month]==="JAN"?"JANUARY":months[month]==="FEB"?"FEBRUARY":months[month]==="MAR"?"MARCH":months[month]==="APR"?"APRIL":months[month]==="JUN"?"JUNE":months[month]==="JUL"?"JULY":months[month]==="MAY"?"MAY":"AUGUST"}</h1></button>
    <div className="calendar">{calendarCells.map((d,index)=>d===null?<span className="calendar-blank" key={`blank-${index}`}/>:<button key={d} onClick={()=>importedDays.has(d)&&setSelectedDay(d)} className={`${d===today?"today":""} ${index%7===0?"sun":""} ${index%7===6?"sat":""} ${isCurrentMonth&&today!==null&&d>today&&!importedDays.has(d)?"future":""} ${importedDays.has(d)&&!completedByDay.has(d)?"has-words":""} ${selectedDay===d&&importedDays.has(d)?"selected-day":""}`}>
      {completedByDay.has(d)?<img src={completedByDay.get(d)!.src}/>:d}
    </button>)}</div>
    <div className="start-area">{selectedDay&&importedDays.has(selectedDay)?<><small>{month+1}/{selectedDay} · {importedWords.filter(w=>w.date===`${year}-${String(month+1).padStart(2,"0")}-${String(selectedDay).padStart(2,"0")}`).length} words</small><button onClick={()=>selectDay(selectedDay)}>Start</button></>:<small>Select a highlighted date</small>}</div>
    {importError&&<div className="import-error">{importError}<button onClick={()=>setImportError("")}>×</button></div>}
    {importedWords.length>0&&<div className="import-success">✓ {importedWords.length} words imported</div>}
    <input ref={fileRef} className="file-input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>readFile(e.target.files?.[0])}/><button className="import-card" onClick={()=>fileRef.current?.click()}><span><Upload/></span><label>Import Word List<small>Excel / CSV</small></label><ChevronRight/></button>
    {monthPicker&&<div className="month-overlay" onClick={()=>setMonthPicker(false)}><section className="month-picker" onClick={e=>e.stopPropagation()}><header><button onClick={()=>setYear(year-1)}>‹</button><b>{year}</b><button onClick={()=>setYear(year+1)}>›</button></header><div>{months.map((m,i)=><button key={m} className={month===i?"selected":""} onClick={()=>{setMonth(i);setMonthPicker(false)}}><b>{i+1}</b><small>{m}</small></button>)}</div></section></div>}
  </div>
}

function Import({allocation,setAllocation,go,words,fileName,confirm}:{allocation:Allocation;setAllocation:(a:Allocation)=>void;go:(s:Screen)=>void;words:ImportedWord[];fileName:string;confirm:(w:ImportedWord[])=>void}){
  const dated=words.filter(w=>w.date); const undated=words.filter(w=>!w.date);
  const groups=Object.entries(dated.reduce<Record<string,number>>((a,w)=>{a[w.date!]=(a[w.date!]||0)+1;return a},{})).sort(([a],[b])=>a.localeCompare(b));
  const displayDate=(s:string)=>new Date(`${s}T00:00:00`).toLocaleDateString("en-US",{month:"short",day:"numeric"});
  const finalized=()=>{
    if(!undated.length)return words;
    const start=new Date("2026-08-12T00:00:00");
    return words.map((w,i)=>{if(w.date)return w;if(allocation==="today")return {...w,date:"2026-08-12"};const d=new Date(start);d.setDate(start.getDate()+Math.floor(i/10));return {...w,date:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`}})
  };
  return <div className="page import"><div className="nav"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h2>Confirm Import</h2><i/></div>
    <img className="hero" src={hornet}/><h1>{words.length} Words Found</h1><p className="file-name">{fileName}</p>{undated.length>0&&<><h3>{undated.length} Undated Words</h3>
    <div className="segments"><button className={allocation==="today"?"on":""} onClick={()=>setAllocation("today")}>All Today</button><button className={allocation==="split"?"on":""} onClick={()=>setAllocation("split")}>Split by Day</button></div>
    {allocation==="split"&&<div className="split-row"><button>10 words / day</button><button>Starting Aug 12</button></div>}</>}
    <div className="date-rows">{groups.slice(0,3).map(([date,count])=><p key={date}>▣　 {displayDate(date)} <b>{count}</b></p>)}{groups.length>3&&<p className="more">▣　 {groups.length-3} More Dates <b>{groups.slice(3).reduce((n,[,c])=>n+c,0)}</b></p>}</div>
    <h3 className="preview-title">Word Preview</h3><div className="preview">{words.slice(0,3).map(w=><p key={w.en}>{w.en} <span>{w.zh}</span><i>{w.date?displayDate(w.date):"No date"}</i></p>)}</div>
    <p className="note">ⓘ　{undated.length?"Blank dates follow your choice above":"All dates were recognized successfully"}</p><button className="primary" onClick={()=>confirm(finalized())}>Import {words.length} Words</button>
  </div>
}

function Characters({selected,setSelected,go}:{selected:number;setSelected:(n:number)=>void;go:(s:Screen)=>void}){
  return <div className="page character"><div className="top"><IconButton onClick={()=>go("home")} label="Close"><X/></IconButton><span>?</span></div><h1>Choose your character ✨</h1>
    <div className="characters">{characters.map((c,i)=><button key={c.name} className={selected===i?"chosen":""} onClick={()=>setSelected(i)} aria-label={c.name}><img src={c.src}/>{selected===i&&<b>✓</b>}</button>)}</div>
    <button className="next" onClick={()=>go("quiz")}>→</button>
  </div>
}

function Quiz({mode,chosen,answer,setAnswer,question,setQuestion,go,reviewOnly,round,words,attempts,setAttempts}:{mode:Mode;chosen:Character;answer:string;setAnswer:(s:string)=>void;question:number;setQuestion:(n:number)=>void;go:(s:Screen)=>void;reviewOnly:boolean;round:number;words:ImportedWord[];attempts:Attempt[];setAttempts:(a:Attempt[])=>void}){
  const quizWords=reviewOnly?words.slice(0,3):words;
  const total=Math.max(quizWords.length,1); const current=quizWords[Math.min(question-1,quizWords.length-1)]??{en:"beautiful",zh:"美丽的",date:null};
  const [advancing,setAdvancing]=useState(false);
  const voiceRef=useRef<SpeechSynthesisVoice|null>(null);
  useEffect(()=>{
    const lockVoice=()=>{
      if(voiceRef.current)return;
      const voices=speechSynthesis.getVoices();
      voiceRef.current=voices.find(v=>v.lang==="en-US"&&/Samantha|Ava|Alex/i.test(v.name))
        ||voices.find(v=>v.lang==="en-US"&&/Google US English/i.test(v.name))
        ||voices.find(v=>v.lang==="en-US")
        ||voices.find(v=>v.lang.startsWith("en"))
        ||null;
    };
    lockVoice();
    speechSynthesis.addEventListener("voiceschanged",lockVoice);
    return()=>speechSynthesis.removeEventListener("voiceschanged",lockVoice);
  },[]);
  const prompt=mode==="Dictation"?"Listen and type the word":mode==="C → E"?"Translate into English":"Translate into Chinese";
  const expected=mode==="E → C"?current.zh:current.en;
  const normalized=(s:string)=>s.trim().toLocaleLowerCase().replace(/[，,；;]/g,";").replace(/\s+/g," ");
  const submit=(skip=false)=>{
    if(advancing||(!skip&&!answer.trim()))return;
    setAdvancing(true);
    const typed=skip?"":answer;
    const correct=!skip&&(mode==="E → C"?normalized(expected).split(";").some(x=>normalized(typed)===x):normalized(typed)===normalized(expected));
    const next=[...attempts,{...current,answer:typed,status:skip?"skip":correct?"correct":"wrong",round} as Attempt];
    setAttempts(next);setAnswer("");
    if(question>=total){go("result");return}
    setQuestion(question+1);
    window.setTimeout(()=>setAdvancing(false),180);
  };
  const speak=()=>{
    const utterance=new SpeechSynthesisUtterance(current.en);utterance.lang="en-US";utterance.rate=.78;utterance.pitch=1;
    if(voiceRef.current)utterance.voice=voiceRef.current;
    speechSynthesis.cancel();
    window.setTimeout(()=>speechSynthesis.speak(utterance),60);
  };
  return <div className="page quiz"><div className="quiz-head"><IconButton onClick={()=>go("home")} label="Close"><X/></IconButton><b>{question} / {total}</b><i/></div><div className="progress"><i style={{width:`${question/total*100}%`}}/></div>
    <h1>{prompt}</h1><div className="prompt-area">{mode==="Dictation"?<button className="sound" onClick={speak}><Volume2/></button>:<strong>{mode==="C → E"?current.zh:current.en}</strong>}<img src={chosen.src}/></div>
    {mode==="Dictation"&&<small>Tap to replay</small>}<div className="answer"><input autoFocus autoCapitalize="none" autoCorrect="off" spellCheck={false} inputMode={mode==="E → C"?"text":"latin"} value={answer} onChange={e=>setAnswer(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&answer.trim())submit(false)}} placeholder={mode==="E → C"?"Enter the Chinese meaning":"Type the English word"}/><i/></div>
    <div className="quiz-actions"><button disabled={advancing} onClick={()=>submit(true)}>I don’t know</button><button className="primary" disabled={!answer.trim()||advancing} onClick={()=>submit(false)}>Submit</button></div>
  </div>
}

function Result({chosen,mode,go,retry,done,attempts,round}:{chosen:Character;mode:Mode;go:(s:Screen)=>void;retry:()=>void;done:()=>void;attempts:Attempt[];round:number}){
  const shown=attempts.filter(a=>a.round===round);
  const correct=shown.filter(a=>a.status==="correct").length, wrong=shown.filter(a=>a.status==="wrong").length, skipped=shown.filter(a=>a.status==="skip").length;
  const accuracy=shown.length?correct/shown.length*100:0;
  return <div className="page result"><div className="nav"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h2>{mode} Result<small>Aug 12, 2026</small></h2><IconButton onClick={()=>go("journal")} label="Search"><Search/></IconButton></div>
    <div className="summary"><img src={chosen.src}/><p>{shown.length} words · {correct} correct · {wrong} wrong · {skipped} unanswered<br/>Accuracy　<strong>{accuracy.toFixed(1)}%</strong></p></div><h3>{round===1?"All Words":`Review Round ${round-1}`}</h3>
    <div className="word-list">{shown.map((w,i)=><div className="word" key={`${w.en}-${i}`}><b className={w.status}>{w.status==="correct"?"✓":w.status==="wrong"?"×":"−"}</b><p><strong>{i+1}. {w.en}</strong>{w.status==="correct"?<small>Correct</small>:<small>Your answer: <em>{w.answer||"—"}</em><i>Correct: <u>{mode==="E → C"?w.zh:w.en}</u></i></small>}</p></div>)}</div>
    <div className="result-actions"><button onClick={done}>Done</button>{wrong+skipped>0&&<button className="primary" onClick={retry}>Practice {wrong+skipped} Again</button>}</div>
  </div>
}

function Journal({go,sessions}:{go:(s:Screen)=>void;sessions:Session[]}){return <div className="page journal"><div className="year"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h1>2026</h1><Search/></div>
  {sessions.length===0?<div className="empty-journal">No practice records yet</div>:sessions.map(session=>{
    const roundNumbers=Array.from(new Set(session.attempts.map(a=>a.round??1))).sort((a,b)=>a-b);
    const firstRound=session.attempts.filter(a=>(a.round??1)===1);
    const correct=firstRound.filter(a=>a.status==="correct").length;
    const wrong=firstRound.filter(a=>a.status==="wrong").length;
    const skipped=firstRound.filter(a=>a.status==="skip").length;
    const accuracy=firstRound.length?correct/firstRound.length*100:0;
    const date=new Date(`${session.date}T00:00:00`);
    const missed=firstRound.filter(a=>a.status==="wrong");
    return <article key={session.id}><img src={session.character.src}/><header><b>{date.toLocaleDateString("en-US",{month:"short",day:"numeric"})}</b><small>{date.toLocaleDateString("en-US",{weekday:"short"})}</small><i>{session.mode}</i></header><ChevronRight/><p>
      Today I practiced {firstRound.length} words: <em>{firstRound.map(a=>a.en).join(", ")}</em>.<br/>
      I got {correct} correct, {wrong} wrong, and left {skipped} unanswered.<br/>
      Accuracy: {accuracy.toFixed(1)}%.<br/>
      {missed.length>0&&<>I misspelled {missed.map((a,i)=><span key={`${a.en}-${i}`}>{i>0&&(i===missed.length-1?" and ":", ")}“{a.en}” as <strong>“{a.answer}”</strong></span>)}.</>}
      {missed.length===0&&skipped===0&&<>Great work — every answer was correct.</>}
      {missed.length===0&&skipped>0&&<>There {skipped===1?"was":"were"} {skipped} unanswered {skipped===1?"word":"words"}.</>}
      {roundNumbers.filter(r=>r>1).map(r=>{const items=session.attempts.filter(a=>(a.round??1)===r);const ok=items.filter(a=>a.status==="correct").length;const bad=items.filter(a=>a.status!=="correct");return <span className="round-summary" key={r}><br/><b>Review {r-1}:</b> {items.length} words, {ok} correct, {bad.length} still incorrect.{bad.length>0&&<> Remaining: {bad.map((a,i)=><span key={`${a.en}-${i}`}>{i>0?", ":""}<em>{a.en}</em> (“{a.answer||"—"}”)</span>)}.</>}</span>})}
    </p></article>
  })}</div>}

export default App;
