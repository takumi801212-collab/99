const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5
pres.author = "LAWSON STORE 100";
pres.title  = "デカ盛りチャレンジ タグライン ＆ POPコピー設計";

const INK="16181C", PAPER="FFFFFF", SURF="F1F1EC", SURF2="E7E7E0",
      POP="D8232A", HL="F7DE3A", MUTED="6E7076", RULE="D6D6CE",
      DINK="242830", DMUTED="9BA0A8", OK="1F7A4D";
const FJ="Meiryo";
const W=13.333, M=0.7, CW=W-M*2;

const R = () => pres.ShapeType.rect;

function dark(){ const s=pres.addSlide(); s.background={color:INK}; return s; }
function light(){ const s=pres.addSlide(); s.background={color:PAPER}; return s; }

function chip(s, text, x, y, o={}){
  const bg=o.bg||POP, fg=o.fg||PAPER, size=o.size||10, h=o.h||0.3;
  const w=o.w || (0.34 + [...text].length*0.158);
  s.addShape(R(), {x,y,w,h, fill:{color:bg}, line:{color:bg,width:0}});
  s.addText(text, {x,y,w,h, fontSize:size, bold:true, color:fg, align:"center",
                   valign:"middle", fontFace:FJ, margin:0});
  return w;
}
function title(s, text, o={}){
  s.addText(text, {x:M, y:o.y||1.02, w:CW, h:0.75, fontSize:o.size||30, bold:true,
                   color:o.color||INK, fontFace:FJ, margin:0, valign:"middle"});
}
function eyebrow(s, text, o={}){
  s.addText(text, {x:M, y:o.y||0.6, w:CW, h:0.3, fontSize:11, bold:true,
                   color:o.color||POP, fontFace:FJ, margin:0, charSpacing:2, valign:"middle"});
}
function card(s,x,y,w,h,o={}){
  s.addShape(R(), {x,y,w,h, fill:{color:o.bg||SURF}, line:{color:o.line||(o.bg||SURF), width:o.lw||0}});
}

/* ───────────────────────── 1. TITLE ───────────────────────── */
{
  const s=dark();
  chip(s,"確定版",M,0.85,{size:11});
  s.addText("ローソンストア100  ／  デカ盛りチャレンジ",
    {x:M+1.05,y:0.85,w:7,h:0.3,fontSize:11,bold:true,color:DMUTED,fontFace:FJ,margin:0,valign:"middle",charSpacing:1});
  s.addText("『デカ盛りチャレンジ』\nタグライン ＆ POPコピー設計",
    {x:M,y:1.75,w:11.4,h:2.1,fontSize:40,bold:true,color:PAPER,fontFace:FJ,margin:0,lineSpacing:52});
  s.addText("通常回 ／ 主戦場は店頭POP ／ 単独開催 ／「お値段そのまま」使用不可",
    {x:M,y:4.1,w:11.4,h:0.4,fontSize:15,color:DMUTED,fontFace:FJ,margin:0});

  const items=[["¥157","ばくだんおにぎり"],["¥289","トリプルたまごサンド"],["¥538","ダブルロースカツカレー"]];
  items.forEach(([p,n],i)=>{
    const x=M+i*3.9;
    s.addShape(R(),{x,y:5.15,w:1.15,h:0.45,fill:{color:POP},line:{color:POP,width:0}});
    s.addText(p,{x,y:5.15,w:1.15,h:0.45,fontSize:17,bold:true,color:PAPER,align:"center",valign:"middle",fontFace:FJ,margin:0});
    s.addText(n,{x:x+1.3,y:5.15,w:2.5,h:0.45,fontSize:11,color:DMUTED,valign:"middle",fontFace:FJ,margin:0});
  });
  s.addText("2026.08.20",{x:M,y:6.55,w:5,h:0.3,fontSize:11,color:MUTED,fontFace:FJ,margin:0});
  s.addNotes("確定版。前提4条件（通常回／POP主戦場／単独開催／お値段そのまま不可）のもとでの設計。");
}

/* ───────────────────────── 2. TAGLINE ───────────────────────── */
{
  const s=dark();
  eyebrow(s,"確定タグライン",{y:0.85});
  s.addText([
    {text:"値上げには、",options:{color:PAPER}},
    {text:"\n",options:{}},
    {text:"デカ盛り",options:{color:HL}},
    {text:"で。",options:{color:PAPER}},
  ],{x:M,y:1.55,w:8.2,h:2.6,fontSize:60,bold:true,fontFace:FJ,margin:0,lineSpacing:76});

  s.addText("ローソンストア100　デカ盛りチャレンジ",
    {x:M,y:4.25,w:8,h:0.35,fontSize:14,color:DMUTED,fontFace:FJ,margin:0});

  card(s,M,5.05,11.93,1.55,{bg:DINK});
  chip(s,"社内で止まった場合の即差し替え案",M+0.4,5.32,{size:10,bg:HL,fg:INK});
  s.addText("物価高には、デカ盛りで。",
    {x:M+0.4,y:5.72,w:5.4,h:0.6,fontSize:26,bold:true,color:PAPER,fontFace:FJ,margin:0,valign:"middle"});
  s.addText("文字数・改行位置・組み方はすべて同一。\n入稿直前でも1語の差し替えで対応でき、三層設計は一切変わらない。",
    {x:M+6.2,y:5.7,w:5.3,h:0.7,fontSize:12,color:DMUTED,fontFace:FJ,margin:0,valign:"middle",lineSpacing:19});
  s.addNotes("本命は据え置き。保険案は版下を作り直さずに1語だけ差し替えられる。");
}

/* ───────────────────────── 3. 6つの回答 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"01 ／ 前提の確定");
  title(s,"6つの回答が決めたこと");
  const data=[
    ["1","通常回","実績語を封印。毎回まったく同じ言葉で回す方針が正解に"],
    ["2","「値上げ」表現を推したい","本命を据え置き。説得材料と即差し替え案をセットで用意"],
    ["3","盛りすぎと並走しない","親会社との描き分けが不要に。制約が1つ消え、言葉を強くできる"],
    ["4","主戦場はPOP","設計の中心がタグライン1本から「三層のPOP体系」へ移動"],
    ["5","1個買うと1個と並走しない","上位コピー不要。デカ盛り単体に全予算・全面積を集中できる"],
    ["6","「お値段そのまま」不可","最重要。本命の選択が「好み」から「必然」に変わった"],
  ];
  const cw=3.81, ch=1.78, gx=0.25, gy=0.28;
  data.forEach(([n,q,a],i)=>{
    const col=i%3, row=Math.floor(i/3);
    const x=M+col*(cw+gx), y=2.05+row*(ch+gy);
    const key = n==="6";
    card(s,x,y,cw,ch,{bg:key?POP:SURF});
    chip(s,n,x+0.32,y+0.3,{w:0.34,size:11,bg:key?PAPER:INK,fg:key?POP:PAPER});
    s.addText(q,{x:x+0.82,y:y+0.28,w:cw-1.14,h:0.36,fontSize:13.5,bold:true,
                 color:key?PAPER:INK,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(a,{x:x+0.32,y:y+0.78,w:cw-0.64,h:0.82,fontSize:11,
                 color:key?PAPER:MUTED,fontFace:FJ,margin:0,lineSpacing:17});
  });
  s.addText("6番の制約が、本命タグラインの選択を消去法ではなく必然として確定させた。以降の設計はすべてこの1点から導かれている。",
    {x:M,y:6.25,w:11.93,h:0.4,fontSize:12.5,bold:true,color:INK,fontFace:FJ,margin:0});
  s.addNotes("6番が本提案の骨子を決定づけた。");
}

/* ───────────────────────── 4. 骨子 ───────────────────────── */
{
  const s=dark();
  eyebrow(s,"02 ／ 骨子",{y:0.6,color:HL});
  title(s,"お得さを名乗れない以上、語れるものは2つしかない",{color:PAPER,size:27});

  const bw=5.72, bh=2.35;
  [["姿勢","タグラインが担当","値上げには、デカ盛りで。"],
   ["目の前の事実","商品札が担当","カツ、2枚。　¥538"]].forEach(([h,r,ex],i)=>{
    const x=M+i*(bw+0.49);
    card(s,x,2.25,bw,bh,{bg:DINK});
    chip(s,r,x+0.42,2.55,{size:10,bg:i?PAPER:HL,fg:INK});
    s.addText(h,{x:x+0.42,y:3.0,w:bw-0.84,h:0.55,fontSize:26,bold:true,color:PAPER,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(ex,{x:x+0.42,y:3.7,w:bw-0.84,h:0.5,fontSize:16,bold:true,color:HL,fontFace:FJ,margin:0,valign:"middle"});
  });

  s.addText("これは制約ではなく、他社との差が最も開く条件です。",
    {x:M,y:5.1,w:11.93,h:0.4,fontSize:15,bold:true,color:PAPER,fontFace:FJ,margin:0});
  s.addText("競合3社はいずれも「お値段そのまま＋◯%増量」という比較の言葉で戦っている。比較が使えないLS100は、必然的に比較しない言葉で戦うことになる。\nそしてその戦い方は、元の価格が最初から低い店でしか成立しない。",
    {x:M,y:5.6,w:11.93,h:0.95,fontSize:13,color:DMUTED,fontFace:FJ,margin:0,lineSpacing:22});
  s.addNotes("6番の制約が、本命案を消去法ではなく必然として確定させた。");
}

/* ───────────────────────── 5. 競合 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"03 ／ 競合");
  title(s,"3社はすべて同じ構文で戦っている");
  const comp=[["ファミリーマート","お値段そのまま\n45%増量作戦","45%","2021〜／作戦"],
              ["ローソン","盛りすぎ\nチャレンジ","約50%","2023.02〜／チャレンジ"],
              ["セブン‐イレブン","感謝盛り／\n人気商品増量祭","50%+","創業感謝祭／祭"]];
  const cw=3.81, gx=0.25;
  comp.forEach(([n,t,pc,meta],i)=>{
    const x=M+i*(cw+gx);
    card(s,x,2.0,cw,2.45,{bg:SURF});
    s.addText(n,{x:x+0.34,y:2.22,w:cw-0.68,h:0.32,fontSize:12,bold:true,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(pc,{x:x+0.34,y:2.6,w:cw-0.68,h:0.8,fontSize:38,bold:true,color:POP,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(t,{x:x+0.34,y:3.45,w:cw-0.68,h:0.6,fontSize:13,bold:true,color:INK,fontFace:FJ,margin:0,lineSpacing:19});
    s.addText(meta,{x:x+0.34,y:4.05,w:cw-0.68,h:0.28,fontSize:10,color:MUTED,fontFace:FJ,margin:0});
  });
  card(s,M,4.78,11.93,0.86,{bg:INK});
  s.addText([
    {text:"お値段そのまま",options:{color:HL}},
    {text:"　＋　",options:{color:MUTED}},
    {text:"◯◯%増量",options:{color:HL}},
    {text:"　＋　",options:{color:MUTED}},
    {text:"作戦／チャレンジ／祭",options:{color:HL}},
  ],{x:M,y:4.78,w:11.93,h:0.86,fontSize:19,bold:true,align:"center",valign:"middle",fontFace:FJ,margin:0});
  s.addText("いずれも自社通常品との相対値＝比較の言葉。今回LS100はこの構文を一切使えない。",
    {x:M,y:5.9,w:11.93,h:0.4,fontSize:13,bold:true,color:INK,fontFace:FJ,margin:0});
  s.addText("出典：ファミリーマート／ローソン／セブン‐イレブン 各社ニュースリリース",
    {x:M,y:6.62,w:11.93,h:0.32,fontSize:10,color:MUTED,fontFace:FJ,margin:0});
  s.addNotes("出典：各社ニュースリリース。");
}

/* ───────────────────────── 6. 客層 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"04 ／ 客層");
  title(s,"「◯%お得」では動かない客層である");
  const stats=[["53.1%","節約パパ・ママ","これ一品で食卓が片づく\n＝ 家計の安心"],
               ["20.9%","食欲旺盛ワーカー","これ一個で足りる\n＝ 満腹の担保"],
               ["約4割","50代以上","コンビニ業界平均より\n5〜10歳高い"]];
  const cw=3.81,gx=0.25;
  stats.forEach(([v,l,d],i)=>{
    const x=M+i*(cw+gx);
    card(s,x,2.05,cw,2.6,{bg:i===2?SURF2:SURF});
    s.addText(v,{x:x+0.34,y:2.3,w:cw-0.68,h:0.95,fontSize:46,bold:true,color:i===2?INK:POP,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(l,{x:x+0.34,y:3.3,w:cw-0.68,h:0.36,fontSize:15,bold:true,color:INK,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(d,{x:x+0.34,y:3.75,w:cw-0.68,h:0.7,fontSize:11.5,color:MUTED,fontFace:FJ,margin:0,lineSpacing:18});
  });
  card(s,M,5.0,11.93,0.95,{bg:INK});
  s.addText("2大セグメントに共通するのは「足りる」という感覚であって、「◯%お得」という計算ではない。",
    {x:M+0.45,y:5.0,w:11.03,h:0.95,fontSize:15,bold:true,color:PAPER,fontFace:FJ,margin:0,valign:"middle"});
  s.addText("出典：流通ニュース（ローソンストア100 顧客構成／小栗知義社長インタビュー）",
    {x:M,y:6.35,w:11.93,h:0.32,fontSize:10,color:MUTED,fontFace:FJ,margin:0});
  s.addNotes("出典：流通ニュース。節約パパ・ママ53.1%、食欲旺盛ワーカー20.9%。");
}

/* ───────────────────────── 7. POP三層 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"05 ／ 主戦場");
  title(s,"POP三層設計");
  const tiers=[
    ["第1層","入口・什器上部・大判","3–5m ／ 0.5秒","値上げには、デカ盛りで。","姿勢の宣言。商品を1つも説明せず、足を止めさせて次の層へ引き渡す。","不変"],
    ["第2層","売場什器・棚上ボード","1–2m ／ 1–2秒","その値段で、そのデカさ。","価格と量の橋渡し。視線を商品と値札の往復に誘導する。比較語ゼロ。","不変"],
    ["第3層","商品札・プライスカード脇","30cm ／ 2–3秒","カツ、2枚。　¥538","最後のひと押し。形容詞を使わず、数えられる事実か身体の感覚だけ。","毎回"],
  ];
  const bh=1.36, gy=0.18;
  tiers.forEach(([t,place,dist,copy,role,upd],i)=>{
    const y=1.92+i*(bh+gy);
    card(s,M,y,11.93,bh,{bg:i===2?SURF2:SURF});
    s.addText(t,{x:M+0.34,y:y+0.22,w:1.5,h:0.34,fontSize:15,bold:true,color:POP,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(place+"\n"+dist,{x:M+0.34,y:y+0.6,w:2.3,h:0.62,fontSize:10,color:MUTED,fontFace:FJ,margin:0,lineSpacing:16});
    s.addText(copy,{x:M+3.0,y:y+0.24,w:4.75,h:0.95,fontSize:i===2?20:22,bold:true,color:INK,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(role,{x:M+8.0,y:y+0.26,w:3.05,h:0.9,fontSize:10.5,color:MUTED,fontFace:FJ,margin:0,lineSpacing:16,valign:"middle"});
    chip(s,upd,M+11.15,y+0.55,{w:0.62,size:10,bg:upd==="毎回"?POP:INK});
  });
  s.addText("版を作り直すのは第3層の商品札だけ。制作コストと入稿工数が毎回ほぼ商品札だけで済み、回を重ねるほど絵柄が資産になる。",
    {x:M,y:6.55,w:11.93,h:0.4,fontSize:12.5,bold:true,color:INK,fontFace:FJ,margin:0});
  s.addNotes("層をまたいで同じことを言わない。これが面積あたりの効率を最大化する。");
}

/* ───────────── 7.5 ロックアップ（タグライン×企画名） ───────────── */
{
  const s=light();
  eyebrow(s,"06 ／ ロックアップ");
  title(s,"タグラインと企画名は、同じ扱いで並べない");
  const pw=5.72, ph=3.0, py=1.92;

  // NG
  card(s,M,py,pw,ph,{bg:SURF});
  chip(s,"NG",M+0.4,py+0.3,{w:0.56,size:11,bg:POP});
  s.addText("同じ級数・同じ書体で3行並べる",
    {x:M+1.12,y:py+0.3,w:pw-1.5,h:0.3,fontSize:11,bold:true,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
  s.addText("値上げには、\nデカ盛りで。\nデカ盛りチャレンジ",
    {x:M+0.4,y:py+0.98,w:pw-0.8,h:1.3,fontSize:20,bold:true,color:INK,fontFace:FJ,margin:0,lineSpacing:29});
  s.addText("「デカ盛り」が同じ扱いで2回出る。読み手には、同じことを2回言っているようにしか見えない。",
    {x:M+0.4,y:py+2.36,w:pw-0.8,h:0.5,fontSize:11,color:POP,fontFace:FJ,margin:0,lineSpacing:17});

  // OK
  card(s,M+pw+0.49,py,pw,ph,{bg:SURF2});
  const ox=M+pw+0.49;
  chip(s,"OK",ox+0.4,py+0.3,{w:0.56,size:11,bg:INK});
  s.addText("タグラインは「言葉」、企画名は「情報」",
    {x:ox+1.12,y:py+0.3,w:pw-1.5,h:0.3,fontSize:11,bold:true,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
  s.addText("値上げには、\nデカ盛りで。",
    {x:ox+0.4,y:py+0.92,w:pw-0.8,h:1.1,fontSize:26,bold:true,color:INK,fontFace:FJ,margin:0,lineSpacing:37});
  s.addText("デカ盛りチャレンジ　5/13（水）〜5/26（火）　ローソンストア100",
    {x:ox+0.4,y:py+2.08,w:pw-0.8,h:0.3,fontSize:9.5,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
  s.addText("企画名は級数を1/4以下、色は無彩色、位置は下端。開催期間と同じ「情報」の階層に落とす。",
    {x:ox+0.4,y:py+2.42,w:pw-0.8,h:0.5,fontSize:11,color:INK,fontFace:FJ,margin:0,lineSpacing:17});

  // 代替案
  card(s,M,5.18,11.93,1.3,{bg:INK});
  chip(s,"重複を完全に消す場合の代替案",M+0.42,5.42,{size:10,bg:HL,fg:INK});
  s.addText("値上げには、デカ盛りチャレンジで。",
    {x:M+0.42,y:5.82,w:5.6,h:0.5,fontSize:22,bold:true,color:PAPER,fontFace:FJ,margin:0,valign:"middle"});
  s.addText("タグラインが企画名を内包するため、企画名の別掲そのものが不要になる。\nただし16字に伸び、語尾が重くなるぶん切れ味は落ちる。第2候補。",
    {x:M+6.6,y:5.5,w:4.9,h:0.8,fontSize:11.5,color:DMUTED,fontFace:FJ,margin:0,lineSpacing:18,valign:"middle"});

  s.addText("「デカ盛り」の連呼自体は店頭では悪くない。問題になるのは、同じ級数・同じ書体で縦に並べたときだけ。",
    {x:M,y:6.66,w:11.93,h:0.4,fontSize:12.5,bold:true,color:INK,fontFace:FJ,margin:0});
  s.addNotes("タグラインと企画名ロゴの関係を規定するスライド。級数比は4:1以上を目安に。");
}

/* ───────────────────────── 8. 3つの型 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"07 ／ 商品札");
  title(s,"商品札は3つの型で量産する");
  const types=[
    ["型A","数える ── 最優先","〔数えられる事実〕。＋ 価格","カツ、2枚。\nたまご、3段。\nメンチ、3枚。","数字は事実なので法務リスクがない。「2枚」は「たっぷり」より必ず具体的に効く。",POP],
    ["型B","身体 ── 数えられないとき","〔身体感覚〕。＋ 価格","片手に、のらない。\n持つと、重い。\nスプーンが、小さく見える。","おにぎりやプリンなど個数で語れない商品用。必ず実物で成立を確認してから使う。",INK],
    ["型C","価格直球 ── 迷ったらこれ","これで、〔価格〕円。","これで、171円。\nこれで、157円。","事実ゼロ、価格のみ。法務確認が最も簡単で、LS100の価格帯では単体でも十分に強い。",MUTED],
  ];
  const cw=3.81,gx=0.25;
  types.forEach(([id,label,form,ex,why,c],i)=>{
    const x=M+i*(cw+gx);
    card(s,x,2.0,cw,4.35,{bg:SURF});
    chip(s,id,x+0.34,2.28,{w:0.62,size:11,bg:c});
    s.addText(label,{x:x+1.06,y:2.28,w:cw-1.4,h:0.3,fontSize:11,bold:true,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(form,{x:x+0.34,y:2.78,w:cw-0.68,h:0.34,fontSize:11,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(ex,{x:x+0.34,y:3.25,w:cw-0.68,h:1.5,fontSize:16,bold:true,color:INK,fontFace:FJ,margin:0,lineSpacing:28});
    s.addText(why,{x:x+0.34,y:4.95,w:cw-0.68,h:1.2,fontSize:11,color:MUTED,fontFace:FJ,margin:0,lineSpacing:18});
  });
  s.addText("形容詞（大きい・たっぷり・ボリューム満点）は全型で禁止。主観であり、比較の含意が生まれやすく、しかも実物より弱い。",
    {x:M,y:6.55,w:11.93,h:0.4,fontSize:12.5,bold:true,color:POP,fontFace:FJ,margin:0});
  s.addNotes("担当者が誰でも同じ品質で書けるよう型を固定する。優先順位は A → B → C。");
}

/* ───────────────────────── 9. 適用例 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"08 ／ 適用例");
  title(s,"直近ラインナップへの適用");
  const items=[
    ["ダブルロースカツカレー","¥538","カツ、2枚。","A",""],
    ["トリプルたまごサンド","¥289","たまご、3段。","A",""],
    ["トリプルメンチカツサンド","¥298","メンチ、3枚。","A",""],
    ["大きな白身魚フライ＆\nハンバーグ弁当","¥538","主役が、2つ。","A",""],
    ["ばくだんおにぎり\n和風ツナマヨ","¥157","片手に、のらない。","B","※実物で要確認"],
    ["ロールケーキ\nりんご＆カスタード","¥171","これで、171円。","C",""],
  ];
  const cw=3.81,ch=2.02,gx=0.25,gy=0.24;
  items.forEach(([n,p,c,t,flag],i)=>{
    const col=i%3,row=Math.floor(i/3);
    const x=M+col*(cw+gx), y=2.0+row*(ch+gy);
    card(s,x,y,cw,ch,{bg:SURF});
    s.addText(n,{x:x+0.32,y:y+0.22,w:cw-1.0,h:0.55,fontSize:10.5,color:MUTED,fontFace:FJ,margin:0,lineSpacing:15});
    chip(s,t,x+cw-0.62,y+0.22,{w:0.34,size:10,bg:SURF2,fg:INK});
    s.addShape(R(),{x:x+0.32,y:y+0.86,w:0.78,h:0.36,fill:{color:POP},line:{color:POP,width:0}});
    s.addText(p,{x:x+0.32,y:y+0.86,w:0.78,h:0.36,fontSize:13,bold:true,color:PAPER,align:"center",valign:"middle",fontFace:FJ,margin:0});
    s.addText(c,{x:x+0.32,y:y+1.32,w:cw-0.64,h:0.42,fontSize:17,bold:true,color:INK,fontFace:FJ,margin:0,valign:"middle"});
    if(flag) s.addText(flag,{x:x+cw-1.42,y:y+1.38,w:1.15,h:0.3,fontSize:9,color:POP,fontFace:FJ,margin:0,align:"right",valign:"middle"});
  });
  s.addText("価格は過去回の税込価格の仮置き。実施回の確定価格に差し替えのうえ、全商品ぶんを型に沿って書き切ります。",
    {x:M,y:6.5,w:11.93,h:0.4,fontSize:12,color:MUTED,fontFace:FJ,margin:0});
  s.addNotes("対象商品リストと確定価格をいただければ全商品ぶん作成します。");
}

/* ───────────────────────── 10. 禁止語 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"09 ／ 法務");
  title(s,"禁止語と言い換え");
  s.addText("制作会社と共有すれば、上がってきた版下の一次チェックがその場で終わります。",
    {x:M,y:1.82,w:11.93,h:0.32,fontSize:12,color:MUTED,fontFace:FJ,margin:0});
  const rows=[
    ["お値段そのまま／価格据え置き","価格の実数をそのまま出す"],
    ["◯%増量／1.5倍／大幅増量","2枚・3段など数えられる事実"],
    ["通常品より／前回より","比較せず、目の前の事実だけ書く"],
    ["お得／実質値下げ／激安","価格を出し、判断は客に渡す"],
    ["業界最大／日本一／No.1","使わない（通常回のため封印）"],
    ["大きい／たっぷり／ボリューム満点","型A・型B・型C"],
  ];
  chip(s,"使わない",M,2.28,{size:10});
  chip(s,"言い換える",M+6.5,2.28,{size:10,bg:INK});
  const rh=0.62, rg=0.1;
  rows.forEach(([bad,good],i)=>{
    const y=2.78+i*(rh+rg);
    card(s,M,y,6.1,rh,{bg:SURF});
    s.addText(bad,{x:M+0.28,y,w:5.54,h:rh,fontSize:12.5,bold:true,color:POP,fontFace:FJ,margin:0,valign:"middle"});
    s.addText("→",{x:M+6.16,y,w:0.32,h:rh,fontSize:13,color:MUTED,align:"center",valign:"middle",fontFace:FJ,margin:0});
    card(s,M+6.5,y,5.43,rh,{bg:SURF2});
    s.addText(good,{x:M+6.78,y,w:4.87,h:rh,fontSize:12.5,bold:true,color:INK,fontFace:FJ,margin:0,valign:"middle"});
  });
  s.addNotes("法務に確認いただきたいのは「値上げには」を『当店は値上げしない』と読ませないかの1点のみ。第2層以下は比較語を含まないため確認不要と考える。");
}

/* ───────────────────────── 11. 版下ルール ───────────────────────── */
{
  const s=light();
  eyebrow(s,"10 ／ 版下");
  title(s,"POP制作のルール");
  const rules=[
    ["改行位置を固定する","「値上げには、／デカ盛りで。」の2行組を正とする。1行組は横に間延びし、遠見での視認速度が落ちる。読点2つが呼吸をつくるので縦組みでもそのまま成立する。"],
    ["書体は極太ゴシック一択","細い書体や明朝は企画内容と矛盾する。言葉の意味と文字の太さを一致させる。"],
    ["企画色を固定し、毎回変えない","通常回を毎回同じ色で打つことでしか、店頭での記号化は起きない。回ごとの色替えは資産形成を毎回ゼロに戻す。"],
    ["コピー単体のPOPを作らない","必ず価格が同じ視野に入るよう配置する。比較訴求が使えない以上、説得力の最大値は価格の実数にある。"],
    ["第3層は1商品1メッセージ","2つ以上の情報を載せると、30cm・2秒では両方読まれない。"],
  ];
  const rh=0.9, rg=0.12;
  rules.forEach(([h,d],i)=>{
    const y=1.98+i*(rh+rg);
    card(s,M,y,11.93,rh,{bg:SURF});
    chip(s,String(i+1),M+0.32,y+0.27,{w:0.36,h:0.36,size:12,bg:INK});
    s.addText(h,{x:M+0.94,y,w:3.5,h:rh,fontSize:14,bold:true,color:INK,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(d,{x:M+4.68,y,w:6.95,h:rh,fontSize:11.5,color:MUTED,fontFace:FJ,margin:0,lineSpacing:18,valign:"middle"});
  });
  s.addNotes("制作会社への発注仕様としてそのまま使える。");
}

/* ───────────────────────── 12. 社内説得 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"11 ／ 社内");
  title(s,"「値上げ」表現を通す4つの論拠");
  s.addText("会議で出る反論はほぼ1つ ──「自社広告でネガティブワードを使うのか」。",
    {x:M,y:1.82,w:11.93,h:0.32,fontSize:12,color:MUTED,fontFace:FJ,margin:0});
  const args=[
    ["主語は「世の中の値上げ」","特定企業の非難ではない",false],
    ["「値上げしない」とは言っていない","言っているのは「値上げに量で応える」という姿勢だけ。実態と矛盾しない",false],
    ["この企画はお得さを名乗れない","法務確認により「お値段そのまま」が使えない。名乗れない以上、語れるのは姿勢だけ",false],
    ["新しい主張ではない","直近リリース見出し「値上げが続く今こそ応援！」の主語と述語を残しただけ。社としてすでに言っている言葉を、POPに載る長さに研いだもの",true],
  ];
  const rh=0.9, rg=0.14;
  args.forEach(([h,d,key],i)=>{
    const y=2.22+i*(rh+rg);
    card(s,M,y,11.93,rh,{bg:key?POP:SURF});
    chip(s,String(i+1),M+0.32,y+0.27,{w:0.36,h:0.36,size:12,bg:key?PAPER:INK,fg:key?POP:PAPER});
    s.addText(h,{x:M+0.94,y,w:4.1,h:rh,fontSize:14,bold:true,color:key?PAPER:INK,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(d,{x:M+5.15,y,w:6.5,h:rh,fontSize:11.5,color:key?PAPER:MUTED,fontFace:FJ,margin:0,lineSpacing:17,valign:"middle"});
  });
  s.addText("4が最も効きます。「新しく踏み込む案」ではなく「すでに社が言っていることの圧縮」だと示せれば、論点は表現の可否から文字数の話に降ります。",
    {x:M,y:6.55,w:11.93,h:0.4,fontSize:12.5,bold:true,color:INK,fontFace:FJ,margin:0});
  s.addNotes("通らなかった場合は「物価高には、デカ盛りで。」へ。版下を作り直さずに1語だけ差し替えられる。");
}

/* ───────────────────────── 13. 運用 ───────────────────────── */
{
  const s=light();
  eyebrow(s,"12 ／ 運用");
  title(s,"通常回だからこそ、変えない");
  const rows=[
    ["第1層タグライン","不変","一字も変えない。数年かけて記号化する"],
    ["第2層ボード","不変","版を作り直さない"],
    ["企画色・書体・改行","不変","回ごとに変えると資産形成が毎回ゼロに戻る"],
    ["第3層 商品札","毎回","型A→B→Cの優先順で機械的に生成"],
    ["実績語（過去最大 等）","封印","通常回では使わない。特別回のために温存する"],
  ];
  const rh=0.68, rg=0.10;
  rows.forEach(([el,freq,pol],i)=>{
    const y=1.95+i*(rh+rg);
    card(s,M,y,11.93,rh,{bg:SURF});
    s.addText(el,{x:M+0.32,y,w:3.3,h:rh,fontSize:13.5,bold:true,color:INK,fontFace:FJ,margin:0,valign:"middle"});
    chip(s,freq,M+3.8,y+0.17,{w:0.72,size:10.5,bg:freq==="毎回"?POP:(freq==="封印"?MUTED:INK)});
    s.addText(pol,{x:M+4.85,y,w:6.76,h:rh,fontSize:12,color:MUTED,fontFace:FJ,margin:0,valign:"middle"});
  });
  card(s,M,6.05,11.93,0.92,{bg:INK});
  s.addText("3回目の掲出から、客は言葉を読まずに絵柄で企画を認識しはじめる。",
    {x:M+0.45,y:6.05,w:11.03,h:0.92,fontSize:15,bold:true,color:HL,fontFace:FJ,margin:0,valign:"middle"});
  s.addNotes("定期企画で最も効くのは、強い言葉を毎回考え直すことではなく、一度決めて絶対に変えないこと。");
}

/* ───────────────────────── 14. 次の一手 ───────────────────────── */
{
  const s=dark();
  eyebrow(s,"13 ／ 次の一手",{color:HL});
  title(s,"この設計を動かすために必要なもの",{color:PAPER});
  const needs=[
    ["今回の対象商品リストと確定価格","いただければ全商品ぶんの第3層コピーを型に沿って書き切ります"],
    ["「値上げには」の法務確認","絞った1点のみ。結果次第で保険案に切り替えます"],
    ["型Bを使う商品の実物確認","成立しなければ型Cに落とします"],
    ["POP什器のサイズ・掲出面数","三層それぞれの版下サイズを詰めるのに必要です"],
  ];
  const rh=0.8, rg=0.16;
  needs.forEach(([h,d],i)=>{
    const y=2.15+i*(rh+rg);
    chip(s,String(i+1).padStart(2,"0"),M,y+0.19,{w:0.46,h:0.36,size:11,bg:HL,fg:INK});
    s.addText(h,{x:M+0.75,y:y+0.02,w:5.0,h:0.75,fontSize:15,bold:true,color:PAPER,fontFace:FJ,margin:0,valign:"middle"});
    s.addText(d,{x:M+5.95,y:y+0.02,w:5.98,h:0.75,fontSize:11.5,color:DMUTED,fontFace:FJ,margin:0,lineSpacing:17,valign:"middle"});
  });
  s.addText([
    {text:"値上げには、",options:{color:PAPER}},
    {text:"デカ盛り",options:{color:HL}},
    {text:"で。",options:{color:PAPER}},
  ],{x:M,y:6.1,w:8,h:0.7,fontSize:28,bold:true,fontFace:FJ,margin:0,valign:"middle"});
  s.addNotes("次のアクション一覧。");
}

pres.writeFile({fileName:"/tmp/claude-0/-home-user-99/7852433a-1c88-51dc-a33f-9cbf867a2df4/scratchpad/dekamori-tagline.pptx"})
  .then(f=>console.log("written:",f));
