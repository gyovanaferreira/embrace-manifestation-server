const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, BorderStyle, ShadingType, PageBreak } = require('docx');
const { Resend } = require('resend');

const app = express();
app.use(cors());
app.use(express.json({limit:'50mb'}));

const ANTHROPIC_KEY = process.env.ANTHROPIC_KEY;
const RESEND_KEY = process.env.RESEND_KEY;
const resend = new Resend(RESEND_KEY);

// Colors
const NAVY="1C2B3A", TEAL="2A7F7F", ORANGE="C8541A", GREY="6B7280";
const TEAL_BG="EAF4F4", ORANGE_BG="FDF0E8", GOLD="C8A882";

// ── HEALTH CHECK ──────────────────────────────────────────────────────────────
app.get('/', (req, res) => res.send('Embrace Manifestation Server running. Key loaded: ' + (ANTHROPIC_KEY ? 'YES' : 'NO')));

// ── GENERATE + EMAIL ──────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  const { name, area, goal, email } = req.body;
  if(!name||!area||!goal||!email) return res.status(400).json({error:'Missing fields'});

  try {
    // CALL 1: Days 1-11
    const txt1 = await callClaude(buildPrompt(name, area, goal, 1));
    // CALL 2: Days 12-21
    const txt2 = await callClaude(buildPrompt(name, area, goal, 2));

    const p1 = parseJSON(txt1);
    const p2 = parseJSON(txt2);
    const days = (p1.days||[]).concat(p2.days||[]);
    days.forEach((d,i) => d.day = i+1);

    const plan = { intro: p1.intro, days };

    // Generate Word doc
    const docBuffer = await buildDocx(name, area, goal, plan);

    // Send email with attachment
    await resend.emails.send({
      from: 'Embrace Manifestation <onboarding@resend.dev>',
      to: [email],
      subject: `${name}'s 21-Day ${area} Identity Reset — Your Action Map`,
      html: buildEmailHtml(name, area, goal),
      attachments: [{
        filename: `${name}-${area}-Action-Map.docx`,
        content: docBuffer.toString('base64')
      }]
    });

    res.json({ success: true, days: plan.days.length });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── CLAUDE API ────────────────────────────────────────────────────────────────
async function callClaude(prompt) {
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 8000,
      messages: [{ role: 'user', content: prompt }]
    })
  });
  const data = await res.json();
  if(data.error) throw new Error(data.error.message);
  return data.content&&data.content[0]&&data.content[0].text||'';
}

function parseJSON(txt) {
  const clean = txt.replace(/^```json\s*/,'').replace(/```\s*$/,'').trim();
  const start = clean.indexOf('{');
  const end = clean.lastIndexOf('}');
  if(start===-1||end===-1) throw new Error('JSON incomplete — response was cut off');
  return JSON.parse(clean.slice(start,end+1));
}

// ── PROMPT ────────────────────────────────────────────────────────────────────
function buildPrompt(name, area, goal, callNum) {
  const lessonAssignment = callNum===1
    ? 'Day 1 uses F1+F2+F3+F4+F5+F6+D1 (all foundations plus D1). Days 2-11 use D2 through D11 in order.'
    : 'Days 12-21 use D12 through D21 in order, one lesson per day.';
  const dayInstruction = callNum===1
    ? 'Generate ONLY days 1 through 11. Start day numbering at 1.'
    : 'Generate ONLY days 12 through 21. Start day numbering at 12.';
  const videoInstruction = callNum===1
    ? 'Day 1 videos: all 6 foundations + "Day 1 — Recognizing the Life Loop". Days 2-11: just their daily video title.'
    : 'Each day: just their daily video title e.g. "Day 12 — Upgrading Your Inner Dialogue".';

  return `You are the AI engine for the Embrace Manifestation Action Map by Gyovana (@embrace.manifestation).

Create a personalized action map for ${name} whose goal is "${goal}" in the area of ${area}.

${lessonAssignment}
${dayInstruction}

LESSON LIBRARY:
F1-Why Your Life Repeats Itself: Exercise — write the Life Loop: situation that repeats, usual reaction, outcome, then the rule underneath. Find 2 examples from last 6 months.
F2-How Your Brain Filters Reality: Exercise — set phone alarm, ask what opportunity you have been filtering out.
F3-The Predictive Brain: Exercise — identify one automatic prediction your brain makes. Write: my brain predicts X which makes me do Y.
F4-The Nervous System Safety Gate: Exercise — before any key moment pause and breathe 3 times. Say aloud: it is safe for me to have more in this area.
F5-Fear Mode vs Heart Mode: Exercise — notice if you are in Fear Mode or Heart Mode. Before key moments do 3 breaths to shift.
F6-How the Brain Changes Through Evidence: Exercise — write 3 pieces of evidence that contradict your limiting belief.
D1-Recognizing the Life Loop: Exercise — write Life Loop: situation, reaction, outcome, rule underneath. Find 2 specific examples.
D2-Seeing Your Current Identity Clearly: Exercise — complete 5 times: when it comes to [area] my brain believes I am someone who... Connect to the loop.
D3-The Thoughts That Define You: Exercise — set 2 phone alarms (11am, 4pm), write last area-related thought. Track all thoughts. Circle most frequent.
D4-Emotional Baselines: Exercise — before every key moment notice what emotion is in your body. Name it precisely. Compare before and after regulation.
D5-Decisions Reveal Identity: Exercise — track every decision point today. Write: situation, decision, automatic or conscious?
D6-Interrupting Autopilot: Exercise — choose 3 moments to pause before responding. Take breath, name the automatic reaction.
D7-The First Identity Shift: Exercise — do regulation first. Choose one real action future self would take. Write before and after.
D8-Designing Your Future Self: Exercise — write Future Identity. Answer: if this area worked as I want, what person would I be? Write 5 sentences.
D9-Acting Before You Feel Ready: Exercise — identify one postponed action. Write prediction before. Take it. Write: prediction vs what happened.
D10-Nervous System Expansion: Exercise — take one stretching action. Before: write body sensations. Take action. After: write sensations. Do regulation.
D11-Creating Identity Evidence: Exercise — look for 3+ moments today where behavior reflects who you are becoming. Write each immediately.
D12-Upgrading Your Inner Dialogue: Exercise — write the most repeated sentence about this area. Write its effect. Write a believable upgrade sentence.
D13-Changing Your Environment: Exercise — identify environment most triggering old pattern. Make one specific change. Create one intentional positive context.
D14-The Power of Micro Decisions: Exercise — track every micro decision today. Before each: which choice does the woman I am becoming make?
D15-Embodying the Future You: Exercise — before every key moment: pause, slow breath, soften shoulders, ask how future self would carry herself.
D16-Identity Through Action: Exercise — ask: what one daily action does the woman I am becoming do consistently? Do that action today.
D17-Trusting the Process: Exercise — write the doubt exactly. Open Day 1 journal and read it. Write 3 things actually different now.
D18-Recognizing Early Reality Shifts: Exercise — notice what you see differently today. Find one moment where your reaction was different from Day 1.
D19-Becoming the Creator: Exercise — write: in what area do I feel most like a creator? Take one proactive action the old you would have waited for.
D20-Your New Internal Standard: Exercise — write standard: From now on I will no longer [list]. From now on I will always [list].
D21-Locking the New Identity: Exercise — reread Day 1 journal, write before and after. Write final identity statement. Design 3-part daily practice.

RULES:
1. ${videoInstruction}
2. Personalize EVERY element to the goal "${goal}" and area ${area}. No generic content.
3. exercise_steps: 3-4 steps, concrete, specific to her goal, include short examples in parentheses.
4. journal_example: 2-3 sentences, specific to her goal and area.
5. Voice: warm, direct, second person "you". Grade 5 reading level.

Respond ONLY with valid JSON, no markdown:
{"intro":"2 warm sentences for ${name} about her journey","days":[{"day":1,"about":"2 sentences personalized to her goal","videos":["video title"],"exercise_title":"title","exercise_intro":"1 sentence","exercise_steps":["step 1 (example)","step 2","step 3"],"journal_prompt":"personalized question","journal_example":"2-3 sentence specific example"}]}`;
}

// ── WORD DOC BUILDER ──────────────────────────────────────────────────────────
async function buildDocx(name, area, goal, plan) {
  const SZ = {footer:16,label:18,body:22,italic:22,dayNum:28,h2:36,h1:44,coverSub:30,coverMain:52};

  function sp(b=100,a=80){ return new Paragraph({spacing:{before:b,after:a},children:[new TextRun("")]}); }
  function hr(col=TEAL){ return new Paragraph({spacing:{before:0,after:0},border:{bottom:{style:BorderStyle.SINGLE,size:6,color:col}}}); }
  function lbl(text,col=TEAL){ return new Paragraph({spacing:{before:260,after:80},children:[new TextRun({text:text.toUpperCase(),bold:true,size:SZ.label,color:col,font:"Calibri",characterSpacing:60})]}); }
  function body(text,col=NAVY,italic=false){ return new Paragraph({spacing:{before:80,after:80},children:[new TextRun({text,size:SZ.body,color:col,italic,font:"Calibri"})]}); }
  function dash(text){ return new Paragraph({spacing:{before:80,after:80},children:[new TextRun({text:"—  ",bold:true,color:TEAL,size:SZ.body,font:"Calibri"}),new TextRun({text,size:SZ.body,color:NAVY,font:"Calibri"})]}); }
  function box(text,bgCol,bdrCol){ return new Paragraph({spacing:{before:100,after:100},shading:{fill:bgCol,type:ShadingType.CLEAR},border:{top:{style:BorderStyle.SINGLE,size:3,color:bdrCol},bottom:{style:BorderStyle.SINGLE,size:3,color:bdrCol},left:{style:BorderStyle.THICK,size:14,color:bdrCol},right:{style:BorderStyle.SINGLE,size:3,color:bdrCol}},children:[new TextRun({text,size:SZ.body,color:NAVY,italic:true,font:"Georgia"})]}); }
  function exBox(text){ return new Paragraph({spacing:{before:80,after:100},shading:{fill:"FAF7F2",type:ShadingType.CLEAR},border:{top:{style:BorderStyle.SINGLE,size:2,color:"D1D5DB"},bottom:{style:BorderStyle.SINGLE,size:2,color:"D1D5DB"},left:{style:BorderStyle.SINGLE,size:6,color:GOLD},right:{style:BorderStyle.SINGLE,size:2,color:"D1D5DB"}},children:[new TextRun({text:"Example: "+text,size:SZ.label,color:GREY,italic:true,font:"Calibri"})]}); }
  function wline(){ return [new Paragraph({spacing:{before:60,after:0},border:{bottom:{style:BorderStyle.SINGLE,size:3,color:"D1D5DB"}},children:[new TextRun({text:" ",size:SZ.body})]}),sp(80,0)]; }
  function chkItem(text){ return new Paragraph({spacing:{before:80,after:80},children:[new TextRun({text:"☐  ",bold:true,color:TEAL,size:SZ.body,font:"Calibri"}),new TextRun({text,size:SZ.body,color:NAVY,font:"Calibri"})]}); }

  function actionItem(text, n) {
    const lines = text.split("\n");
    return lines.map((ln,i) => new Paragraph({spacing:{before:i===0?120:40,after:40},children:[
      i===0 ? new TextRun({text:`${n}.  `,bold:true,color:ORANGE,size:SZ.body,font:"Calibri"}) : new TextRun({text:"     ",size:SZ.body,font:"Calibri"}),
      new TextRun({text:ln.replace(/^   /,""),size:SZ.body,color:NAVY,font:"Calibri"})
    ]}));
  }

  const children = [];

  // COVER
  children.push(sp(500,0));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:80},children:[new TextRun({text:"@embrace.manifestation",size:SZ.label,color:TEAL,font:"Calibri",characterSpacing:80})]}));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:60},children:[new TextRun({text:`${name}'s`,size:SZ.coverSub,color:GREY,font:"Georgia",italic:true})]}));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:60},children:[new TextRun({text:`21-Day ${area} Identity Reset`,size:SZ.coverMain,color:NAVY,bold:true,font:"Georgia"})]}));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:200},children:[new TextRun({text:"Your Step-by-Step Action Map",size:SZ.coverSub,color:TEAL,font:"Georgia",italic:true})]}));
  children.push(hr(GOLD));
  children.push(sp(140,0));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:60},children:[new TextRun({text:"GOAL",size:SZ.label,bold:true,color:ORANGE,font:"Calibri",characterSpacing:80})]}));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:300},shading:{fill:ORANGE_BG,type:ShadingType.CLEAR},border:{left:{style:BorderStyle.THICK,size:14,color:ORANGE},right:{style:BorderStyle.THICK,size:14,color:ORANGE},top:{style:BorderStyle.SINGLE,size:3,color:ORANGE},bottom:{style:BorderStyle.SINGLE,size:3,color:ORANGE}},children:[new TextRun({text:`"${goal}"`,size:SZ.h2/2*2,italic:true,color:NAVY,font:"Georgia"})]}));
  children.push(sp(300,0));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:`21 Days  ·  ${area}  ·  Embrace Manifestation`,size:SZ.label,color:GREY,font:"Calibri"})]}));

  // HOW TO USE
  children.push(new Paragraph({children:[new PageBreak()]}));
  children.push(new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:0,after:140},children:[new TextRun({text:"How This Works",bold:true,size:SZ.h1,color:NAVY,font:"Georgia"})]}));
  children.push(hr(TEAL));
  children.push(sp(120,0));
  children.push(body("Each day has the same simple structure. Follow the steps in order. Tick each box when done. Then move on.",GREY,true));
  children.push(sp(160,0));
  children.push(lbl("Every day — four steps",NAVY));
  children.push(chkItem("Step 1 — Watch your lesson videos"));
  children.push(chkItem("Step 2 — Do the exercise from the lesson"));
  children.push(chkItem("Step 3 — Nervous system regulation (5 minutes)"));
  children.push(chkItem("Step 4 — Journal reflection + evidence logging"));
  children.push(sp(160,0));
  children.push(body("That is the entire daily practice. It takes about 15 minutes a day. You can do this every single day.",GREY,true));

  // IDENTITY ANCHOR
  children.push(new Paragraph({children:[new PageBreak()]}));
  children.push(new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:0,after:140},children:[new TextRun({text:"Your Identity Anchor",bold:true,size:SZ.h1,color:NAVY,font:"Georgia"})]}));
  children.push(hr(ORANGE));
  children.push(sp(100,0));
  children.push(body("Fill this in on Day 1. Read it every single morning before you begin.",GREY,true));
  children.push(sp(180,0));
  children.push(lbl("I am someone who...",TEAL));
  children.push(...wline(),...wline(),...wline());
  children.push(sp(100,0));
  children.push(lbl("My goal in my own words:",TEAL));
  children.push(...wline(),...wline());
  children.push(sp(100,0));
  children.push(lbl("The old pattern I am releasing:",TEAL));
  children.push(...wline(),...wline());

  // ROADMAP
  children.push(new Paragraph({children:[new PageBreak()]}));
  children.push(new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:0,after:140},children:[new TextRun({text:"Your 21-Day Roadmap",bold:true,size:SZ.h1,color:NAVY,font:"Georgia"})]}));
  children.push(hr(TEAL));
  children.push(sp(80,0));
  children.push(body("Check off each day when you complete it. Every tick is a signal to your brain that you are someone who shows up.",GREY,true));
  children.push(sp(160,0));

  const weeks = [
    {n:1,phase:"Awareness",col:TEAL,bg:TEAL_BG,desc:"See the patterns, identity, thoughts, and decisions running on autopilot."},
    {n:2,phase:"Reprogramming",col:ORANGE,bg:ORANGE_BG,desc:"Design the new identity, act before you feel ready, collect evidence every day."},
    {n:3,phase:"Embodiment",col:"5B8FA8",bg:"EBF2F7",desc:"Stabilize the new identity in your body, behavior, and daily standard."}
  ];

  weeks.forEach(wk => {
    const wkDays = plan.days.filter(d => Math.ceil(d.day/7)===wk.n);
    if(!wkDays.length) return;
    children.push(new Paragraph({spacing:{before:180,after:80},shading:{fill:wk.bg,type:ShadingType.CLEAR},border:{top:{style:BorderStyle.SINGLE,size:3,color:wk.col},bottom:{style:BorderStyle.SINGLE,size:3,color:wk.col},left:{style:BorderStyle.THICK,size:14,color:wk.col},right:{style:BorderStyle.SINGLE,size:3,color:wk.col}},children:[
      new TextRun({text:`WEEK ${wk.n}  —  ${wk.phase.toUpperCase()}    `,bold:true,size:SZ.body,color:wk.col,font:"Calibri",characterSpacing:40}),
      new TextRun({text:wk.desc,size:SZ.body,color:GREY,font:"Calibri",italic:true})
    ]}));
    wkDays.forEach(d => {
      children.push(new Paragraph({spacing:{before:80,after:80},children:[
        new TextRun({text:"☐  ",bold:true,color:wk.col,size:SZ.body,font:"Calibri"}),
        new TextRun({text:`Day ${d.day}  `,size:SZ.body,bold:true,color:NAVY,font:"Calibri"}),
        new TextRun({text:d.exercise_title||d.title||'',size:SZ.body,color:GREY,font:"Calibri"})
      ]}));
    });
  });

  // DAY PAGES
  plan.days.forEach(d => {
    children.push(new Paragraph({children:[new PageBreak()]}));
    const wk = d.day<=7?1:(d.day<=14?2:3);
    const wcol = wk===1?TEAL:(wk===2?ORANGE:"5B8FA8");
    const wbg = wk===1?TEAL_BG:(wk===2?ORANGE_BG:"EBF2F7");
    const wl = wk===1?'Awareness':(wk===2?'Reprogramming':'Embodiment');

    children.push(new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:`DAY ${d.day}  ·  WEEK ${wk}: ${wl.toUpperCase()}  ·  ${name.toUpperCase()}`,size:SZ.footer,color:wcol,font:"Calibri",characterSpacing:60,bold:true})]}));
    children.push(hr(wcol));
    children.push(sp(120,0));
    children.push(new Paragraph({spacing:{before:0,after:40},children:[new TextRun({text:`Day ${d.day}`,size:SZ.dayNum,color:wcol,font:"Georgia",italic:true})]}));
    children.push(new Paragraph({spacing:{before:0,after:80},children:[new TextRun({text:d.exercise_title||'',size:SZ.h2,color:NAVY,bold:true,font:"Georgia"})]}));
    children.push(body(d.about||'',GREY,true));
    children.push(sp(160,0));

    // STEP 1 — Videos
    children.push(lbl(`Step 1 — Watch Today`,wcol));
    children.push(chkItem(d.day===1?"Watch all Foundation videos (F1-F6) + Day 1 video":`Watch your Day ${d.day} video`));
    (d.videos||[]).forEach(v => children.push(body("   — "+v, GREY)));
    children.push(sp(180,0));

    // STEP 2 — Exercise
    children.push(lbl(`Step 2 — ${d.exercise_title||'Your Exercise'}`,wcol));
    children.push(body(d.exercise_intro||'',GREY,true));
    children.push(sp(80,0));
    (d.exercise_steps||[]).forEach((s,i) => children.push(...actionItem(s,i+1)));
    children.push(sp(100,0));
    children.push(lbl("Write your answers here",GREY));
    for(let i=0;i<12;i++) children.push(...wline());
    children.push(sp(180,0));

    // STEP 3 — Regulation
    children.push(lbl("Step 3 — Nervous System Regulation",wcol));
    children.push(chkItem("5 minutes of breathwork or binaural beats. Do this before your journal."));
    children.push(body("Let your body settle. This is non-negotiable — it is what makes everything else stick.",GREY,true));
    children.push(sp(180,0));

    // STEP 4 — Journal
    children.push(lbl("Step 4 — Evening Journal Reflection",wcol));
    children.push(box(d.journal_prompt||'',TEAL_BG,TEAL));
    children.push(sp(60,0));
    if(d.journal_example) children.push(exBox(d.journal_example));
    children.push(sp(100,0));
    children.push(lbl("My reflection",GREY));
    for(let i=0;i<10;i++) children.push(...wline());
    children.push(sp(160,0));

    // Evidence
    children.push(new Paragraph({spacing:{before:100,after:100},shading:{fill:ORANGE_BG,type:ShadingType.CLEAR},border:{top:{style:BorderStyle.SINGLE,size:3,color:ORANGE},bottom:{style:BorderStyle.SINGLE,size:3,color:ORANGE},left:{style:BorderStyle.THICK,size:14,color:ORANGE},right:{style:BorderStyle.SINGLE,size:3,color:ORANGE}},children:[
      new TextRun({text:"Today's evidence for my new identity:  ",bold:true,size:SZ.body,color:ORANGE,font:"Calibri"}),
      new TextRun({text:"One thing I noticed today — however small — that shows I am already shifting.",size:SZ.body,color:NAVY,italic:true,font:"Calibri"})
    ]}));
    for(let i=0;i<3;i++) children.push(...wline());

    children.push(hr("E2DBD2"));
    children.push(new Paragraph({spacing:{before:80,after:0},alignment:AlignmentType.CENTER,children:[new TextRun({text:`Day ${d.day} of 21  ·  ${name}'s Action Map  ·  @embrace.manifestation`,size:SZ.footer,color:GREY,font:"Calibri"})]}));
  });

  // FINAL PAGE
  children.push(new Paragraph({children:[new PageBreak()]}));
  children.push(sp(400,0));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:140},children:[new TextRun({text:"You did it.",size:SZ.h1,bold:true,color:TEAL,font:"Georgia"})]}));
  children.push(hr(GOLD));
  children.push(sp(180,0));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:280},children:[new TextRun({text:"The woman who achieves her goal — she is already here. Keep going.",size:SZ.dayNum,italic:true,color:GREY,font:"Georgia"})]}));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:120},shading:{fill:TEAL_BG,type:ShadingType.CLEAR},border:{top:{style:BorderStyle.SINGLE,size:3,color:TEAL},bottom:{style:BorderStyle.SINGLE,size:3,color:TEAL},left:{style:BorderStyle.THICK,size:14,color:TEAL},right:{style:BorderStyle.SINGLE,size:3,color:TEAL}},children:[new TextRun({text:"Comment TIMELINE on Instagram to join the Embrace Manifestation Academy",size:SZ.body,bold:true,color:TEAL,font:"Calibri"})]}));
  children.push(sp(180,0));
  children.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'"You don\'t manifest what you want — you manifest what you are."',size:SZ.body,italic:true,color:NAVY,font:"Georgia"})]}));

  const doc = new Document({
    styles:{
      default:{document:{run:{font:"Calibri",size:SZ.body,color:NAVY}}},
      paragraphStyles:[{id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:SZ.h1,bold:true,font:"Georgia",color:NAVY},paragraph:{spacing:{before:0,after:140},outlineLevel:0}}]
    },
    sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1440,right:1440,bottom:1440,left:1440}}},children}]
  });

  return await Packer.toBuffer(doc);
}

// ── EMAIL HTML ────────────────────────────────────────────────────────────────
function buildEmailHtml(name, area, goal) {
  return `<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;color:#1C2B3A;">
    <div style="background:#EAF4F4;padding:32px;text-align:center;border-radius:12px;margin-bottom:24px;">
      <p style="font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#2A7F7F;margin:0 0 8px;font-family:monospace;">@embrace.manifestation</p>
      <h1 style="margin:0 0 8px;font-size:24px;">${name}'s 21-Day ${area} Identity Reset</h1>
      <p style="margin:0;font-style:italic;color:#2A7F7F;">Your Step-by-Step Action Map</p>
    </div>
    <p style="font-size:15px;line-height:1.7;">Your personalized 21-day journal is attached to this email as a Word document (.docx).</p>
    <p style="font-size:15px;line-height:1.7;">Open it in Word, Pages, or Google Docs. Print it or use it digitally — it has writing space for every exercise, journal prompt, and evidence log.</p>
    <div style="background:#FDF0E8;border-left:4px solid #C8541A;padding:16px 20px;border-radius:0 8px 8px 0;margin:24px 0;">
      <p style="margin:0;font-style:italic;">"${goal}"</p>
    </div>
    <p style="font-size:15px;line-height:1.7;">This is your map. Follow it day by day. The identity shift is already beginning.</p>
    <div style="background:#EAF4F4;border:1px solid #2A7F7F;border-radius:12px;padding:24px;text-align:center;margin-top:32px;">
      <p style="margin:0 0 12px;font-weight:bold;">Ready to go deeper?</p>
      <a href="https://www.instagram.com/embrace.manifestation" style="background:#2A7F7F;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-size:14px;">Comment TIMELINE to join the Academy →</a>
    </div>
    <p style="text-align:center;font-style:italic;margin-top:24px;font-size:13px;color:#6B7280;">"You don't manifest what you want — you manifest what you are."</p>
  </div>`;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  // Keep server warm — ping every 14 minutes
  setInterval(() => {
    fetch(`http://localhost:${PORT}/`)
      .then(() => console.log('Keep-alive ping'))
      .catch(() => {});
  }, 14 * 60 * 1000);
});
