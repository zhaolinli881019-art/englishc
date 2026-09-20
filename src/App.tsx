import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { ArrowLeft, BarChart3, BookOpen, CalendarDays, ChevronDown, ChevronRight, ChevronUp, Database, Headphones, House, Search, Settings, SlidersHorizontal, Sparkles, Target, TrendingUp, Upload, Volume2, X } from "lucide-react";
import artTourTimeLogo from "./assets/art-tour-time-logo.png";
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

type Screen = "home"|"import"|"character"|"quiz"|"result"|"journal"|"progress"|"settings";
type Mode = "Dictation"|"C → E"|"E → C";
type Allocation = "today"|"split";
type Character = {name:string; src:string};
type Word = {en:string; zh:string; answer:string; status:"correct"|"wrong"|"skip"};
type ImportedWord = {en:string; zh:string; date:string|null};
type Attempt = ImportedWord & {answer:string; status:"correct"|"wrong"|"skip"; round:number};
type Session = {id:number; mode:Mode; character:Character; attempts:Attempt[]; date:string};
type AppSettings = {voiceRate:number;dailyGoal:number;strictChecking:boolean;soundEnabled:boolean};
const defaultSettings:AppSettings={voiceRate:.78,dailyGoal:10,strictChecking:true,soundEnabled:true};

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
const fullMonths = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];

function localDateKey(date=new Date()){
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`;
}

function readStored<T>(key:string,fallback:T):T{
  try{const value=localStorage.getItem(key);return value?JSON.parse(value) as T:fallback}catch{return fallback}
}

export function App(){
  const previewScreen=import.meta.env.DEV?new URLSearchParams(window.location.search).get("preview"):null;
  const previewDate=localDateKey();
  const previewWords:ImportedWord[]=[
    {en:"adventure",zh:"冒险",date:previewDate},{en:"beautiful",zh:"美丽的",date:previewDate},{en:"courage",zh:"勇气",date:previewDate},
    {en:"discover",zh:"发现",date:previewDate},{en:"explore",zh:"探索",date:null},{en:"forest",zh:"森林",date:null},
    {en:"journey",zh:"旅程",date:null},{en:"knowledge",zh:"知识",date:null},{en:"language",zh:"语言",date:null},
    {en:"practice",zh:"练习",date:null},{en:"remember",zh:"记住",date:null},{en:"tomorrow",zh:"明天",date:null}
  ];
  const previewAttempts:Attempt[]=[
    {en:"adventure",zh:"冒险",date:previewDate,answer:"adventure",status:"correct",round:1},{en:"beautiful",zh:"美丽的",date:previewDate,answer:"beatiful",status:"wrong",round:1},
    {en:"courage",zh:"勇气",date:previewDate,answer:"courage",status:"correct",round:1},{en:"discover",zh:"发现",date:previewDate,answer:"discover",status:"correct",round:1},
    {en:"explore",zh:"探索",date:previewDate,answer:"",status:"skip",round:1},{en:"forest",zh:"森林",date:previewDate,answer:"forest",status:"correct",round:1},
    {en:"journey",zh:"旅程",date:previewDate,answer:"jorney",status:"wrong",round:1},{en:"knowledge",zh:"知识",date:previewDate,answer:"knowledge",status:"correct",round:1}
  ];
  const previewSessions:Session[]=[
    {id:1,mode:"Dictation",character:characters[1],attempts:previewAttempts,date:previewDate},
    {id:2,mode:"C → E",character:characters[2],attempts:previewAttempts.map((item,index)=>({...item,status:index===1?"wrong":"correct",answer:index===1?"beautifull":item.en})),date:localDateKey(new Date(new Date().setDate(new Date().getDate()-2)))},
    {id:3,mode:"E → C",character:characters[0],attempts:previewAttempts.map(item=>({...item,status:"correct",answer:item.zh})),date:localDateKey(new Date(new Date().setDate(new Date().getDate()-5)))}
  ];
  const previewTarget=(['import','character','quiz','result','journal','progress','settings'] as Screen[]).includes(previewScreen as Screen)?previewScreen as Screen:"home";
  const [screen,setScreen]=useState<Screen>(previewTarget);
  const [mode,setMode]=useState<Mode>("Dictation");
  const [allocation,setAllocation]=useState<Allocation>("split");
  const [selected,setSelected]=useState(1);
  const [answer,setAnswer]=useState("");
  const [question,setQuestion]=useState(1);
  const [reviewOnly,setReviewOnly]=useState(false);
  const [pendingWords,setPendingWords]=useState<ImportedWord[]>(previewScreen==="import"?previewWords:[]);
  const [importedWords,setImportedWords]=useState<ImportedWord[]>(()=>readStored("englishc.words",[]));
  const [fileName,setFileName]=useState(previewScreen==="import"?"English_Words_September.xlsx":"");
  const [importError,setImportError]=useState("");
  const [activeWords,setActiveWords]=useState<ImportedWord[]>(previewScreen==="character"||previewScreen==="quiz"||previewScreen==="result"?previewWords.map(word=>({...word,date:word.date??previewDate})):[]);
  const [attempts,setAttempts]=useState<Attempt[]>(previewScreen==="result"?previewAttempts:[]);
  const [round,setRound]=useState(1);
  const [sessions,setSessions]=useState<Session[]>(()=>previewScreen==="journal"||previewScreen==="progress"?previewSessions:readStored<Session[]>("englishc.sessions",[]).map(session=>({...session,character:characters.find(character=>character.name===session.character?.name)??characters[0]})));
  const [settings,setSettings]=useState<AppSettings>(()=>readStored("englishc.settings",defaultSettings));
  useEffect(()=>{if(!previewScreen)localStorage.setItem("englishc.words",JSON.stringify(importedWords))},[importedWords,previewScreen]);
  useEffect(()=>{if(!previewScreen)localStorage.setItem("englishc.sessions",JSON.stringify(sessions))},[sessions,previewScreen]);
  useEffect(()=>{if(!previewScreen)localStorage.setItem("englishc.settings",JSON.stringify(settings))},[settings,previewScreen]);
  const chosen=characters[selected];
  const go=(s:Screen)=>setScreen(s);
  const todayKey=localDateKey();
  const todayWords=importedWords.filter(word=>word.date===todayKey).length;
  const finishedWords=sessions.reduce((total,session)=>total+session.attempts.filter(attempt=>(attempt.round??1)===1).length,0);
  const correctWords=sessions.reduce((total,session)=>total+session.attempts.filter(attempt=>(attempt.round??1)===1&&attempt.status==="correct").length,0);
  const average=finishedWords?Math.round(correctWords/finishedWords*100):0;
  return <main className="stage"><div className="desktop-shell">
    <aside className="side-rail">
      <div className="brand"><img className="brand-mark" src={artTourTimeLogo} alt="藝遊時光"/><div><b>Art Tour Time</b><small>Grow a little every day</small></div></div>
      <nav className="main-nav" aria-label="Main navigation">
        <button className={screen==="home"?"active":""} onClick={()=>go("home")}><House/><span>Home</span></button>
        <button className={screen==="character"||screen==="quiz"||screen==="result"?"active":""} onClick={()=>go(activeWords.length?"character":"home")}><BookOpen/><span>Practice</span></button>
        <button className={screen==="journal"?"active":""} onClick={()=>go("journal")}><CalendarDays/><span>Journal</span></button>
        <button className={screen==="progress"?"active":""} onClick={()=>go("progress")}><BarChart3/><span>Progress</span></button>
      </nav>
      <div className="rail-section"><small>TEST MODE</small>{(["Dictation","C → E","E → C"] as Mode[]).map(item=><button key={item} className={mode===item?"active":""} onClick={()=>setMode(item)}><i/>{item}</button>)}</div>
      <div className="rail-quote"><Sparkles/><p>Small steps turn into big progress.</p><span>Keep going!</span></div>
      <button className={`rail-settings ${screen==="settings"?"active":""}`} onClick={()=>go("settings")}><Settings/>Settings</button>
    </aside>
    <section className="content-shell">
      <header className="desktop-header"><div><small>MY LEARNING SPACE</small><b>{screen==="home"?"Home":screen==="journal"?"Learning journal":screen==="character"?"Choose a companion":screen==="quiz"?"Practice time":screen==="result"?"Session results":screen==="progress"?"Learning progress":screen==="settings"?"Settings":"Import words"}</b></div><div className="header-date"><span>{new Date().toLocaleDateString("en-US",{weekday:"long"})}</span><b>{new Date().toLocaleDateString("en-US",{month:"short",day:"numeric"})}</b></div></header>
      <section className="phone">
    {screen==="home"&&<Home mode={mode} setMode={setMode} go={go} importedWords={importedWords} sessions={sessions} selectDate={(date)=>{setActiveWords(importedWords.filter(w=>w.date===date));setQuestion(1);setRound(1);setAnswer("");setAttempts([]);setReviewOnly(false);go("character")}} onFile={(words,name)=>{setPendingWords(words);setFileName(name);setImportError("");go("import")}} importError={importError} setImportError={setImportError}/>} 
    {screen==="import"&&<Import allocation={allocation} setAllocation={setAllocation} go={go} words={pendingWords} fileName={fileName} confirm={(words)=>{setImportedWords(words);setImportError("");go("home")}}/>}
    {screen==="character"&&<Characters selected={selected} setSelected={setSelected} go={go} mode={mode} wordCount={activeWords.length}/>}
    {screen==="quiz"&&<Quiz mode={mode} chosen={chosen} answer={answer} setAnswer={setAnswer} question={question} setQuestion={setQuestion} go={go} reviewOnly={reviewOnly} round={round} words={activeWords} attempts={attempts} setAttempts={setAttempts} settings={settings}/>}
    {screen==="result"&&<Result chosen={chosen} mode={mode} go={go} round={round} attempts={attempts} done={()=>{const date=attempts[0]?.date??localDateKey();setSessions(current=>[{id:Date.now(),mode,character:chosen,attempts:[...attempts],date},...current]);go("journal")}} retry={()=>{const missed=attempts.filter(a=>a.round===round&&a.status!=="correct").map(({en,zh,date})=>({en,zh,date}));setActiveWords(missed);setRound(r=>r+1);setReviewOnly(true);setQuestion(1);setAnswer("");go("quiz")}}/>}
    {screen==="journal"&&<Journal go={go} sessions={sessions} openSession={(session)=>{setAttempts(session.attempts);setRound(1);setMode(session.mode);setSelected(Math.max(0,characters.findIndex(character=>character.name===session.character.name)));go("result")}}/>}
    {screen==="progress"&&<Progress go={go} sessions={sessions}/>}
    {screen==="settings"&&<SettingsPage go={go} settings={settings} setSettings={setSettings}/>}
      </section>
    </section>
    <aside className="insight-panel">
      <div className="profile"><div className="profile-copy"><small>YOUR COMPANION</small><b>{chosen.name}</b></div><img src={chosen.src} alt=""/></div>
      <section className="today-card"><small>TODAY'S JOURNEY</small><h2>{todayWords?<>{todayWords} words<br/>are waiting</>:<>A fresh page<br/>is waiting</>}</h2><p>{todayWords?"Pick today's date on the calendar and start when you're ready.":"Import a word list to begin your next learning adventure."}</p><div className="mini-stats"><span><b>{sessions.length}</b><small>Sessions</small></span><span><b>{average}%</b><small>Accuracy</small></span></div></section>
      <section className="mode-card"><header><span>Current mode</span><BookOpen/></header><b>{mode}</b><p>{mode==="Dictation"?"Listen carefully, then spell the word.":mode==="C → E"?"Turn the Chinese meaning into English.":"Write the Chinese meaning of each word."}</p></section>
      <section className="library-card"><div><small>WORD LIBRARY</small><b>{importedWords.length}</b><span>words ready to practise</span></div><div className="book-stack"><i/><i/><i/></div></section>
      <footer>Designed for calm, focused practice.</footer>
    </aside>
  </div></main>
}

function IconButton({children,onClick,label}:{children:React.ReactNode,onClick:()=>void,label:string}){return <button className="icon" onClick={onClick} aria-label={label}>{children}</button>}

function Home({mode,setMode,go,importedWords,sessions,selectDate,onFile,importError,setImportError}:{mode:Mode;setMode:(m:Mode)=>void;go:(s:Screen)=>void;importedWords:ImportedWord[];sessions:Session[];selectDate:(date:string)=>void;onFile:(w:ImportedWord[],name:string)=>void;importError:string;setImportError:(s:string)=>void}){
  const swipeStart=useRef<{x:number;y:number}|null>(null);
  const finishSwipe=(x:number,y:number)=>{
    const start=swipeStart.current;
    swipeStart.current=null;
    if(!start)return;
    const dx=x-start.x,dy=y-start.y;
    if(dx>55&&Math.abs(dx)>Math.abs(dy)*1.2)go("journal");
  };
  const [menuOpen,setMenuOpen]=useState(false);
  const [monthPicker,setMonthPicker]=useState(false);
  const currentDate=new Date();
  const [year,setYear]=useState(()=>currentDate.getFullYear());
  const [month,setMonth]=useState(()=>currentDate.getMonth());
  const isCurrentMonth=year===currentDate.getFullYear()&&month===currentDate.getMonth();
  const today=isCurrentMonth?currentDate.getDate():null;
  const dayCount=new Date(year,month+1,0).getDate();
  const leadingBlanks=new Date(year,month,1).getDay();
  const calendarCells=[...Array.from({length:leadingBlanks},()=>null),...Array.from({length:dayCount},(_,i)=>i+1)];
  const availableDays=Array.from(new Set(importedWords.filter(w=>w.date?.startsWith(`${year}-${String(month+1).padStart(2,"0")}`)).map(w=>Number(w.date!.slice(-2))))).sort((a,b)=>a-b);
  const [selectedDay,setSelectedDay]=useState<number|null>(today&&availableDays.includes(today)?today:(availableDays[0]??null));
  useEffect(()=>{
    const firstDate=importedWords.map(w=>w.date).filter((date):date is string=>Boolean(date)).sort()[0];
    if(!firstDate){setSelectedDay(null);return}
    const [nextYear,nextMonth,nextDay]=firstDate.split("-").map(Number);
    setYear(nextYear);
    setMonth(nextMonth-1);
    setSelectedDay(nextDay);
  },[importedWords]);
  useEffect(()=>{
    setSelectedDay(current=>current&&availableDays.includes(current)?current:(availableDays[0]??null));
  },[year,month,importedWords]);
  const fileRef=useRef<HTMLInputElement>(null);
  const moveMonth=(change:number)=>{const next=new Date(year,month+change,1);setYear(next.getFullYear());setMonth(next.getMonth())};
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
  return <div className="page home" onTouchStart={e=>{const touch=e.touches[0];swipeStart.current={x:touch.clientX,y:touch.clientY}}} onTouchMove={e=>{const start=swipeStart.current,touch=e.touches[0];if(start&&touch&&touch.clientX-start.x>12&&Math.abs(touch.clientX-start.x)>Math.abs(touch.clientY-start.y))e.preventDefault()}} onTouchCancel={()=>{swipeStart.current=null}} onTouchEnd={e=>{const touch=e.changedTouches[0];if(touch)finishSwipe(touch.clientX,touch.clientY)}}><div className="top"><button className="home-menu-toggle" onClick={()=>setMenuOpen(!menuOpen)} aria-expanded={menuOpen} aria-label={menuOpen?"Collapse test modes":"Expand test modes"}><House/>{menuOpen?<ChevronUp/>:<ChevronDown/>}</button><IconButton onClick={()=>go("settings")} label="Settings"><Settings/></IconButton></div>
    <div className="desktop-page-heading"><div><small>STUDY PLANNER</small><h1>Make today count.</h1></div><p>Choose a highlighted day to practise, or import a new list to plan ahead.</p></div>
    {menuOpen&&<div className="mode">{(["Dictation","C → E","E → C"] as Mode[]).map(m=><button className={mode===m?"on":""} onClick={()=>{setMode(m);setMenuOpen(false)}} key={m}>{m}</button>)}</div>}
    <section className="planner-card">
      <div className="month-toolbar"><button className="month-step" onClick={()=>moveMonth(-1)} aria-label="Previous month">‹</button><button className="month" onClick={()=>setMonthPicker(true)}><small>{year}</small><h1>{fullMonths[month]}</h1></button><button className="month-step" onClick={()=>moveMonth(1)} aria-label="Next month">›</button></div>
      <div className="weekdays">{["SUN","MON","TUE","WED","THU","FRI","SAT"].map(day=><span key={day}>{day}</span>)}</div>
      <div className="calendar">{calendarCells.map((d,index)=>d===null?<span className="calendar-blank" key={`blank-${index}`}/>:<button key={d} aria-label={`${fullMonths[month]} ${d}`} onClick={()=>importedDays.has(d)&&setSelectedDay(d)} className={`${d===today?"today":""} ${index%7===0?"sun":""} ${index%7===6?"sat":""} ${isCurrentMonth&&today!==null&&d>today&&!importedDays.has(d)?"future":""} ${importedDays.has(d)&&!completedByDay.has(d)?"has-words":""} ${completedByDay.has(d)?"completed-day":""} ${selectedDay===d&&importedDays.has(d)?"selected-day":""}`}>
        {completedByDay.has(d)?<img src={completedByDay.get(d)!.src}/>:<span>{d}</span>}
      </button>)}</div>
      <div className="calendar-legend"><span><i className="legend-today"/>Today</span><span><i className="legend-ready"/>Words ready</span><span><i className="legend-done"/>Completed</span></div>
      <div className="start-area">{selectedDay&&importedDays.has(selectedDay)?<><small>{month+1}/{selectedDay} · {importedWords.filter(w=>w.date===`${year}-${String(month+1).padStart(2,"0")}-${String(selectedDay).padStart(2,"0")}`).length} words</small><button onClick={()=>selectDate(`${year}-${String(month+1).padStart(2,"0")}-${String(selectedDay).padStart(2,"0")}`)}>Start practice <ChevronRight/></button></>:<small>Select a highlighted date to begin</small>}</div>
    </section>
    {importError&&<div className="import-error">{importError}<button onClick={()=>setImportError("")}>×</button></div>}
    <input ref={fileRef} className="file-input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>readFile(e.target.files?.[0])}/><button className="import-card" onClick={()=>fileRef.current?.click()}><span><Upload/></span><label>Import Word List<small>Excel / CSV</small></label><ChevronRight/></button>
    {monthPicker&&<div className="month-overlay" onClick={()=>setMonthPicker(false)}><section className="month-picker" onClick={e=>e.stopPropagation()}><header><button onClick={()=>setYear(year-1)}>‹</button><b>{year}</b><button onClick={()=>setYear(year+1)}>›</button></header><div>{months.map((m,i)=><button key={m} className={month===i?"selected":""} onClick={()=>{setMonth(i);setMonthPicker(false)}}><b>{i+1}</b><small>{m}</small></button>)}</div></section></div>}
  </div>
}

function Import({allocation,setAllocation,go,words,fileName,confirm}:{allocation:Allocation;setAllocation:(a:Allocation)=>void;go:(s:Screen)=>void;words:ImportedWord[];fileName:string;confirm:(w:ImportedWord[])=>void}){
  const dated=words.filter(w=>w.date); const undated=words.filter(w=>!w.date);
  const todayKey=localDateKey();
  const groups=Object.entries(dated.reduce<Record<string,number>>((a,w)=>{a[w.date!]=(a[w.date!]||0)+1;return a},{})).sort(([a],[b])=>a.localeCompare(b));
  const displayDate=(s:string)=>new Date(`${s}T00:00:00`).toLocaleDateString("en-US",{month:"short",day:"numeric"});
  const finalized=()=>{
    if(!undated.length)return words;
    const start=new Date(`${todayKey}T00:00:00`);
    let undatedIndex=0;
    return words.map(w=>{if(w.date)return w;if(allocation==="today")return {...w,date:todayKey};const d=new Date(start);d.setDate(start.getDate()+Math.floor(undatedIndex++/10));return {...w,date:localDateKey(d)}})
  };
  return <div className="page import"><div className="nav"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h2>Confirm Import</h2><i/></div>
    <div className="import-layout"><section className="import-overview">
      <div className="import-hero-block"><img className="hero" src={hornet}/><div><small>UPLOAD READY</small><h1>{words.length} Words Found</h1><p className="file-name">{fileName}</p><p className="import-intro">Review how your words will be added to the learning calendar.</p></div></div>
      {undated.length>0&&<div className="allocation-block"><header><div><small>DATE PLAN</small><h3>{undated.length} Undated Words</h3></div><span>Choose a schedule</span></header>
      <div className="segments"><button className={allocation==="today"?"on":""} onClick={()=>setAllocation("today")}><b>All Today</b><small>Practice in one session</small></button><button className={allocation==="split"?"on":""} onClick={()=>setAllocation("split")}><b>Split by Day</b><small>Build a steady routine</small></button></div>
      {allocation==="split"&&<div className="split-row"><button><small>DAILY GOAL</small><b>10 words / day</b></button><button><small>START DATE</small><b>{displayDate(todayKey)}</b></button></div>}</div>}
      <div className="date-rows">{groups.slice(0,3).map(([date,count])=><p key={date}>▣　 {displayDate(date)} <b>{count}</b></p>)}{groups.length>3&&<p className="more">▣　 {groups.length-3} More Dates <b>{groups.slice(3).reduce((n,[,c])=>n+c,0)}</b></p>}</div>
      <p className="note">ⓘ　{undated.length?"Blank dates follow your choice above":"All dates were recognized successfully"}</p>
    </section><section className="import-preview-panel">
      <header className="preview-head"><div><small>FILE PREVIEW</small><h3>Word Preview</h3></div><span>{words.length} rows</span></header>
      <div className="preview">{words.slice(0,3).map((w,index)=><p key={w.en}><b>{String(index+1).padStart(2,"0")}</b><strong>{w.en}</strong><span>{w.zh}</span><i>{w.date?displayDate(w.date):"No date"}</i></p>)}</div>
      <div className="import-checks"><p><b>✓</b> English and Chinese columns found</p><p><b>✓</b> {dated.length} dated words recognized</p><p className={undated.length?"attention":""}><b>{undated.length?"!":"✓"}</b> {undated.length?`${undated.length} words need a date plan`:"Every word has a study date"}</p></div>
      <button className="primary" onClick={()=>confirm(finalized())}>Import {words.length} Words <ChevronRight/></button>
    </section></div>
  </div>
}

function Characters({selected,setSelected,go,mode,wordCount}:{selected:number;setSelected:(n:number)=>void;go:(s:Screen)=>void;mode:Mode;wordCount:number}){
  const current=characters[selected];
  return <div className="page character"><div className="top"><IconButton onClick={()=>go("home")} label="Close"><X/></IconButton><span>?</span></div><div className="character-intro"><small>ONE LAST STEP</small><h1>Choose your character ✨</h1><p>Your companion will stay with you through this practice session.</p></div>
    <div className="character-layout"><div className="characters">{characters.map((c,i)=><button key={c.name} className={selected===i?"chosen":""} onClick={()=>setSelected(i)} aria-label={c.name}><span className="character-portrait"><img src={c.src}/>{selected===i&&<b>✓</b>}</span><small>{c.name}</small></button>)}</div>
      <aside className="practice-brief"><small>SESSION READY</small><img src={current.src} alt=""/><h2>{current.name} is ready!</h2><dl><div><dt>Mode</dt><dd>{mode}</dd></div><div><dt>Words</dt><dd>{wordCount}</dd></div><div><dt>Companion</dt><dd>{current.name}</dd></div></dl><button onClick={()=>go("quiz")}>Start practice <ChevronRight/></button></aside>
    </div>
    <button className="next" onClick={()=>go("quiz")}><span>Start practice</span><b>→</b></button>
  </div>
}

function Quiz({mode,chosen,answer,setAnswer,question,setQuestion,go,reviewOnly,round,words,attempts,setAttempts,settings}:{mode:Mode;chosen:Character;answer:string;setAnswer:(s:string)=>void;question:number;setQuestion:(n:number)=>void;go:(s:Screen)=>void;reviewOnly:boolean;round:number;words:ImportedWord[];attempts:Attempt[];setAttempts:(a:Attempt[])=>void;settings:AppSettings}){
  const quizWords=words;
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
    const typedValue=normalized(typed),expectedValue=normalized(expected);
    const exact=mode==="E → C"?expectedValue.split(";").some(x=>typedValue===x):typedValue===expectedValue;
    const flexible=mode==="E → C"&&typedValue.length>1&&expectedValue.split(";").some(x=>x.includes(typedValue)||typedValue.includes(x));
    const correct=!skip&&(exact||(!settings.strictChecking&&flexible));
    const next=[...attempts,{...current,answer:typed,status:skip?"skip":correct?"correct":"wrong",round} as Attempt];
    setAttempts(next);setAnswer("");
    if(question>=total){go("result");return}
    setQuestion(question+1);
    window.setTimeout(()=>setAdvancing(false),180);
  };
  const speak=()=>{
    if(!settings.soundEnabled)return;
    const utterance=new SpeechSynthesisUtterance(current.en);utterance.lang="en-US";utterance.rate=settings.voiceRate;utterance.pitch=1;
    if(voiceRef.current)utterance.voice=voiceRef.current;
    speechSynthesis.cancel();
    window.setTimeout(()=>speechSynthesis.speak(utterance),60);
  };
  return <div className="page quiz"><div className="quiz-head"><IconButton onClick={()=>go("home")} label="Close"><X/></IconButton><div><small>{reviewOnly?`Review ${round-1}`:mode}</small><b>{question} / {total}</b></div><i/></div><div className="progress"><i style={{width:`${question/total*100}%`}}/></div>
    <div className="quiz-workspace"><section className="quiz-focus-card"><div className="quiz-kicker"><span>{mode}</span><small>QUESTION {String(question).padStart(2,"0")}</small></div><h1>{prompt}</h1>
      <div className="prompt-area">{mode==="Dictation"?<button className="sound" onClick={speak} aria-label="Play word" disabled={!settings.soundEnabled}><Volume2/><span>{settings.soundEnabled?"Play word":"Sound off"}</span></button>:<strong>{mode==="C → E"?current.zh:current.en}</strong>}<img src={chosen.src}/></div>
      {mode==="Dictation"&&<small className="replay-note">Tap the sound button to replay</small>}
    </section>
    <section className="quiz-answer-panel"><label>Your answer</label><div className="answer"><input autoFocus autoCapitalize="none" autoCorrect="off" spellCheck={false} inputMode={mode==="E → C"?"text":"latin"} value={answer} onChange={e=>setAnswer(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&answer.trim())submit(false)}} placeholder={mode==="E → C"?"Enter the Chinese meaning":"Type the English word"}/><i/></div>
      <div className="quiz-actions"><button disabled={advancing} onClick={()=>submit(true)}>I don’t know</button><button className="primary" disabled={!answer.trim()||advancing} onClick={()=>submit(false)}>Submit <ChevronRight/></button></div>
      <div className="quiz-tips"><span><kbd>Enter</kbd> Submit answer</span><span>Unsure? Skip and review it later.</span></div>
    </section></div>
  </div>
}

function Result({chosen,mode,go,retry,done,attempts,round}:{chosen:Character;mode:Mode;go:(s:Screen)=>void;retry:()=>void;done:()=>void;attempts:Attempt[];round:number}){
  const [filter,setFilter]=useState<"all"|"wrong"|"skip">("all");
  const shown=attempts.filter(a=>a.round===round);
  const correct=shown.filter(a=>a.status==="correct").length, wrong=shown.filter(a=>a.status==="wrong").length, skipped=shown.filter(a=>a.status==="skip").length;
  const accuracy=shown.length?correct/shown.length*100:0;
  const visibleWords=filter==="all"?shown:shown.filter(item=>item.status===filter);
  const resultDate=new Date(`${shown[0]?.date??localDateKey()}T00:00:00`).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});
  return <div className="page result"><div className="nav"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h2>{mode} Result<small>{resultDate}</small></h2><IconButton onClick={()=>go("journal")} label="Search"><Search/></IconButton></div>
    <section className="result-overview"><div className="result-celebration"><img src={chosen.src}/><div><small>{accuracy===100?"PERFECT SESSION":accuracy>=80?"GREAT PROGRESS":"KEEP GROWING"}</small><h1>{accuracy===100?"Amazing work!":accuracy>=80?"Nicely done!":"Every try helps."}</h1><p>{round===1?"Your first-round results are ready.":`Review round ${round-1} is complete.`}</p></div></div><div className="score-ring" style={{background:`conic-gradient(#65ad91 ${accuracy*3.6}deg,#e8ebe7 0)`}}><span><strong>{accuracy.toFixed(0)}%</strong><small>Accuracy</small></span></div></section>
    <div className="result-stats"><span><small>WORDS</small><b>{shown.length}</b></span><span className="stat-correct"><small>CORRECT</small><b>{correct}</b></span><span className="stat-wrong"><small>WRONG</small><b>{wrong}</b></span><span className="stat-skip"><small>UNANSWERED</small><b>{skipped}</b></span></div>
    <div className="result-toolbar"><h3>{round===1?"Word review":`Review Round ${round-1}`}</h3><div><button className={filter==="all"?"active":""} onClick={()=>setFilter("all")}>All <b>{shown.length}</b></button><button className={filter==="wrong"?"active":""} onClick={()=>setFilter("wrong")}>Wrong <b>{wrong}</b></button><button className={filter==="skip"?"active":""} onClick={()=>setFilter("skip")}>Unanswered <b>{skipped}</b></button></div></div>
    <div className="word-list">{visibleWords.length?visibleWords.map((w,i)=><div className="word" key={`${w.en}-${i}`}><b className={w.status}>{w.status==="correct"?"✓":w.status==="wrong"?"×":"−"}</b><p><strong>{shown.indexOf(w)+1}. {w.en}</strong><span className="word-meaning">{w.zh}</span>{w.status==="correct"?<small>Correct</small>:<small>Your answer: <em>{w.answer||"—"}</em><i>Correct: <u>{mode==="E → C"?w.zh:w.en}</u></i></small>}</p></div>):<div className="filter-empty">Nothing to review here ✨</div>}</div>
    <div className="result-actions"><button onClick={done}>Done</button>{wrong+skipped>0&&<button className="primary" onClick={retry}>Practice {wrong+skipped} Again</button>}</div>
  </div>
}

function Journal({go,sessions,openSession}:{go:(s:Screen)=>void;sessions:Session[];openSession:(session:Session)=>void}){
  const [searchOpen,setSearchOpen]=useState(false);
  const [query,setQuery]=useState("");
  const journalYear=sessions[0]?.date?.slice(0,4)??String(new Date().getFullYear());
  const firstAttempts=sessions.flatMap(session=>session.attempts.filter(attempt=>(attempt.round??1)===1));
  const journalCorrect=firstAttempts.filter(attempt=>attempt.status==="correct").length;
  const journalAccuracy=firstAttempts.length?journalCorrect/firstAttempts.length*100:0;
  const studiedDays=new Set(sessions.map(session=>session.date)).size;
  const filteredSessions=sessions.filter(session=>{const text=`${session.date} ${session.mode} ${session.attempts.map(attempt=>`${attempt.en} ${attempt.zh}`).join(" ")}`.toLocaleLowerCase();return text.includes(query.trim().toLocaleLowerCase())});
  return <div className="page journal"><div className="year"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h1>{journalYear}</h1><IconButton onClick={()=>setSearchOpen(current=>!current)} label="Search"><Search/></IconButton></div>
  <div className="journal-heading"><div><small>LEARNING JOURNAL</small><h2>Your words, your progress.</h2><p>Every practice session becomes part of your learning story.</p></div><img src={sessions[0]?.character.src??hornet} alt=""/></div>
  <div className="journal-stats"><span><small>SESSIONS</small><b>{sessions.length}</b></span><span><small>WORDS PRACTISED</small><b>{firstAttempts.length}</b></span><span><small>ACCURACY</small><b>{journalAccuracy.toFixed(0)}%</b></span><span><small>STUDY DAYS</small><b>{studiedDays}</b></span></div>
  {searchOpen&&<div className="journal-search"><Search/><input autoFocus value={query} onChange={event=>setQuery(event.target.value)} placeholder="Search words, dates or modes"/><button onClick={()=>{setQuery("");setSearchOpen(false)}}>×</button></div>}
  <div className="journal-feed-head"><h3>Recent entries</h3><span>{filteredSessions.length} {filteredSessions.length===1?"entry":"entries"}</span></div>
  <section className="journal-feed">{sessions.length===0?<div className="empty-journal">No practice records yet</div>:filteredSessions.length===0?<div className="empty-journal">No matching records</div>:filteredSessions.map(session=>{
    const roundNumbers=Array.from(new Set(session.attempts.map(a=>a.round??1))).sort((a,b)=>a-b);
    const firstRound=session.attempts.filter(a=>(a.round??1)===1);
    const correct=firstRound.filter(a=>a.status==="correct").length;
    const wrong=firstRound.filter(a=>a.status==="wrong").length;
    const skipped=firstRound.filter(a=>a.status==="skip").length;
    const accuracy=firstRound.length?correct/firstRound.length*100:0;
    const date=new Date(`${session.date}T00:00:00`);
    const missed=firstRound.filter(a=>a.status==="wrong");
    return <article key={session.id} role="button" tabIndex={0} onClick={()=>openSession(session)} onKeyDown={event=>{if(event.key==="Enter"||event.key===" ")openSession(session)}}><img src={session.character.src}/><header><b>{date.toLocaleDateString("en-US",{month:"short",day:"numeric"})}</b><small>{date.toLocaleDateString("en-US",{weekday:"short"})}</small><i>{session.mode}</i></header><div className="entry-score"><b>{accuracy.toFixed(0)}%</b><small>accuracy</small></div><ChevronRight/><p>
      Today I practiced {firstRound.length} words: <em>{firstRound.map(a=>a.en).join(", ")}</em>.<br/>
      I got {correct} correct, {wrong} wrong, and left {skipped} unanswered.<br/>
      Accuracy: {accuracy.toFixed(1)}%.<br/>
      {missed.length>0&&<>I misspelled {missed.map((a,i)=><span key={`${a.en}-${i}`}>{i>0&&(i===missed.length-1?" and ":", ")}“{a.en}” as <strong>“{a.answer}”</strong></span>)}.</>}
      {missed.length===0&&skipped===0&&<>Great work — every answer was correct.</>}
      {missed.length===0&&skipped>0&&<>There {skipped===1?"was":"were"} {skipped} unanswered {skipped===1?"word":"words"}.</>}
      {roundNumbers.filter(r=>r>1).map(r=>{const items=session.attempts.filter(a=>(a.round??1)===r);const ok=items.filter(a=>a.status==="correct").length;const bad=items.filter(a=>a.status!=="correct");return <span className="round-summary" key={r}><br/><b>Review {r-1}:</b> {items.length} words, {ok} correct, {bad.length} still incorrect.{bad.length>0&&<> Remaining: {bad.map((a,i)=><span key={`${a.en}-${i}`}>{i>0?", ":""}<em>{a.en}</em> (“{a.answer||"—"}”)</span>)}.</>}</span>})}
    </p></article>
  })}</section></div>
}

function Progress({go,sessions}:{go:(s:Screen)=>void;sessions:Session[]}){
  const attempts=sessions.flatMap(session=>session.attempts.filter(attempt=>(attempt.round??1)===1));
  const correct=attempts.filter(attempt=>attempt.status==="correct").length;
  const accuracy=attempts.length?Math.round(correct/attempts.length*100):0;
  const activeDays=new Set(sessions.map(session=>session.date)).size;
  const activity=Array.from({length:7},(_,index)=>{const date=new Date();date.setDate(date.getDate()-(6-index));const key=localDateKey(date);return {key,label:date.toLocaleDateString("en-US",{weekday:"short"}).slice(0,2),count:sessions.filter(session=>session.date===key).reduce((total,session)=>total+session.attempts.filter(attempt=>(attempt.round??1)===1).length,0)}});
  const activityMax=Math.max(1,...activity.map(day=>day.count));
  const modeStats=(["Dictation","C → E","E → C"] as Mode[]).map(item=>{const items=sessions.filter(session=>session.mode===item).flatMap(session=>session.attempts.filter(attempt=>(attempt.round??1)===1));const ok=items.filter(attempt=>attempt.status==="correct").length;return {mode:item,count:items.length,accuracy:items.length?Math.round(ok/items.length*100):0}});
  const mistakeCounts=new Map<string,{word:string;meaning:string;count:number}>();
  attempts.filter(attempt=>attempt.status!=="correct").forEach(attempt=>{const current=mistakeCounts.get(attempt.en);mistakeCounts.set(attempt.en,{word:attempt.en,meaning:attempt.zh,count:(current?.count??0)+1})});
  const mistakes=Array.from(mistakeCounts.values()).sort((a,b)=>b.count-a.count).slice(0,5);
  return <div className="page progress-page"><div className="nav"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h2>Learning Progress</h2><i/></div>
    <div className="progress-heading"><div><small>YOUR GROWTH</small><h1>See how far you’ve come.</h1><p>Every word you practise adds another leaf to your learning garden.</p></div><TrendingUp/></div>
    <div className="progress-metrics"><section className="progress-score"><div className="score-ring" style={{background:`conic-gradient(#65ad91 ${accuracy*3.6}deg,#e8ebe7 0)`}}><span><strong>{accuracy}%</strong><small>Accuracy</small></span></div><div><small>OVERALL SCORE</small><h2>{accuracy>=90?"Excellent work!":accuracy>=70?"Growing strong":"Keep practising"}</h2><p>Based on {attempts.length} practised words.</p></div></section><span><Target/><small>WORDS PRACTISED</small><b>{attempts.length}</b></span><span><CalendarDays/><small>ACTIVE DAYS</small><b>{activeDays}</b></span><span><Sparkles/><small>SESSIONS</small><b>{sessions.length}</b></span></div>
    <div className="progress-grid"><section className="activity-card"><header><div><small>LAST 7 DAYS</small><h3>Practice activity</h3></div><span>{activity.reduce((total,day)=>total+day.count,0)} words</span></header><div className="activity-chart">{activity.map(day=><div key={day.key}><span><i style={{height:`${Math.max(day.count?18:3,day.count/activityMax*100)}%`}}/>{day.count>0&&<b>{day.count}</b>}</span><small>{day.label}</small></div>)}</div></section>
      <section className="mode-progress-card"><header><small>BY TEST MODE</small><h3>Mode performance</h3></header>{modeStats.map(item=><div className="mode-progress-row" key={item.mode}><div><b>{item.mode}</b><span>{item.count} words</span></div><section><i style={{width:`${item.accuracy}%`}}/></section><strong>{item.accuracy}%</strong></div>)}</section>
      <section className="mistakes-card"><header><div><small>FOCUS NEXT</small><h3>Words to revisit</h3></div><span>{mistakes.length?"Based on recent mistakes":"Nothing needs review"}</span></header>{mistakes.length?<div>{mistakes.map((item,index)=><p key={item.word}><b>{String(index+1).padStart(2,"0")}</b><strong>{item.word}</strong><span>{item.meaning}</span><i>{item.count} {item.count===1?"miss":"misses"}</i></p>)}</div>:<div className="progress-empty">Complete a practice session to see your focus words.</div>}</section>
    </div>
  </div>
}

function SettingsPage({go,settings,setSettings}:{go:(s:Screen)=>void;settings:AppSettings;setSettings:(settings:AppSettings)=>void}){
  const update=(next:Partial<AppSettings>)=>setSettings({...settings,...next});
  return <div className="page settings-page"><div className="nav"><IconButton onClick={()=>go("home")} label="Back"><ArrowLeft/></IconButton><h2>Settings</h2><i/></div>
    <div className="settings-heading"><div><small>MAKE IT YOURS</small><h1>A practice space that fits you.</h1><p>Your preferences are saved automatically on this device.</p></div><Settings/></div>
    <div className="settings-layout"><section className="settings-card"><header><span><Headphones/></span><div><h3>Pronunciation</h3><p>Choose how listening questions sound.</p></div></header><div className="setting-row"><div><b>Word audio</b><small>Enable the pronunciation button</small></div><button className={`setting-switch ${settings.soundEnabled?"on":""}`} onClick={()=>update({soundEnabled:!settings.soundEnabled})} aria-pressed={settings.soundEnabled}><i/></button></div><div className="setting-block"><label>Speaking speed</label><div className="setting-segments"><button className={settings.voiceRate===.65?"active":""} onClick={()=>update({voiceRate:.65})}>Slow</button><button className={settings.voiceRate===.78?"active":""} onClick={()=>update({voiceRate:.78})}>Natural</button><button className={settings.voiceRate===.95?"active":""} onClick={()=>update({voiceRate:.95})}>Quick</button></div></div></section>
      <section className="settings-card"><header><span><Target/></span><div><h3>Daily rhythm</h3><p>Set a comfortable daily word goal.</p></div></header><div className="goal-stepper"><button onClick={()=>update({dailyGoal:Math.max(5,settings.dailyGoal-5)})}>−</button><div><b>{settings.dailyGoal}</b><small>words per day</small></div><button onClick={()=>update({dailyGoal:Math.min(50,settings.dailyGoal+5)})}>+</button></div><div className="goal-scale"><i style={{width:`${settings.dailyGoal/50*100}%`}}/></div></section>
      <section className="settings-card"><header><span><SlidersHorizontal/></span><div><h3>Answer checking</h3><p>Decide how Chinese meanings are matched.</p></div></header><div className="checking-options"><button className={settings.strictChecking?"active":""} onClick={()=>update({strictChecking:true})}><b>Strict</b><small>Match one complete meaning</small></button><button className={!settings.strictChecking?"active":""} onClick={()=>update({strictChecking:false})}><b>Flexible</b><small>Allow partial meaning matches</small></button></div></section>
      <section className="settings-card data-settings"><header><span><Database/></span><div><h3>Your learning data</h3><p>Words, results and settings stay in this browser.</p></div></header><div className="data-note"><b>Stored locally</b><span>No account or cloud connection is required.</span></div><button onClick={()=>go("progress")}>View learning progress <ChevronRight/></button></section>
    </div><p className="settings-saved">✓ Changes are saved automatically</p>
  </div>
}

export default App;
