namespace SmartPro
{
    // =====================================================================
    //  THE PAGE OF THE WEB DASHBOARD
    //  One HTML file with its style and its script. WebDash.Page puts the title in the place of __TITLE__ and the
    //  rows and the design in the place of /*DATA*/null. Inside this C# text every quote is written twice.
    //  The script adds the rows up with the same rules as DashEngine, so the page and the dashboard in Excel
    //  show the same numbers.
    // =====================================================================
    internal static class WebTemplate
    {
        public const string Html = @"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""utf-8"">
<meta name=""viewport"" content=""width=device-width,initial-scale=1"">
<meta name=""generator"" content=""DocBrisk for Excel"">
<title>__TITLE__</title>
<style>
:root{--page:#f1f5f9;--band:#0f172a;--bandText:#ffffff;--card:#ffffff;--text:#0f172a;--muted:#64748b;--grid:#e2e8f0;--on:#dbe6fb;--off:#f3f5f8;--hover:#b9cdf7;--edge:#d0d7e2;--r:10px;--sh:none;--gap:12px;--per:4;--font:'Segoe UI'}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--page);color:var(--text);font:13px/1.4 var(--font),'Segoe UI',system-ui,-apple-system,'Helvetica Neue',Arial,sans-serif}
button,select,input{font:inherit;color:inherit}
header{background:var(--band);color:var(--bandText);padding:14px 24px;display:flex;align-items:center;gap:12px 20px;flex-wrap:wrap}
header.line{border-bottom:1px solid var(--grid)}
header .hd{min-width:0;flex:1 1 320px}
header h1{margin:0;font-size:24px;line-height:1.2;font-weight:700;overflow-wrap:anywhere}
header p{margin:3px 0 0;font-size:12px;opacity:.72}
header img{max-height:52px;max-width:200px;display:block}
.tools{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.tools select,.tools button{background:transparent;color:inherit;border:1px solid currentColor;border-radius:6px;padding:5px 8px;font-size:12px;opacity:.82;cursor:pointer}
.tools select:hover,.tools button:hover,.tools select:focus,.tools button:focus{opacity:1}
.tools option{color:#0f172a;background:#fff}
.chips{max-width:1240px;margin:0 auto;padding:10px 20px 0;display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:12px}
.chips:empty{display:none}
.chip{background:var(--card);border:1px solid var(--edge);border-radius:999px;padding:3px 4px 3px 10px;display:inline-flex;gap:6px;align-items:center;max-width:100%}
.chip span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:60vw}
.chip button{border:0;background:var(--off);border-radius:50%;width:18px;height:18px;line-height:16px;padding:0;cursor:pointer;color:var(--muted);flex:none}
.chips .all{border:0;background:none;color:var(--muted);text-decoration:underline;cursor:pointer;padding:3px 6px}
.wrap{max-width:1240px;margin:0 auto;padding:14px 20px 8px;display:grid;grid-template-columns:210px minmax(0,1fr);gap:var(--gap);align-items:start}
.wide .wrap,.wide .chips,.wide footer{max-width:1680px}
.wrap.p1,.wrap.p3{grid-template-columns:minmax(0,1fr)}
.wrap.p2{grid-template-columns:minmax(0,1fr) 210px}
.wrap.p2 aside{order:2}
aside{position:sticky;top:10px;max-height:calc(100vh - 20px);overflow:auto;padding-bottom:2px;min-width:0}
aside summary{display:none}
#filters{display:grid;gap:var(--gap)}
.wrap.p1 aside{position:static;max-height:none}
.wrap.p1 #filters{display:flex;flex-wrap:wrap}
.sl{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:9px 10px 10px;min-width:0}
.flat .sl,.flat .ch,.flat .ins{border:1px solid var(--edge)}
.sl h3{margin:0 0 6px;font-size:12px;font-weight:700;display:flex;align-items:center;gap:6px}
.sl h3 span{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sl h3 button{border:0;background:none;color:var(--muted);cursor:pointer;font-size:11px;padding:0 2px;text-decoration:underline;visibility:hidden}
.sl.has h3 button{visibility:visible}
.sl input{width:100%;border:1px solid var(--edge);background:var(--card);border-radius:5px;padding:3px 7px;font-size:12px;margin-bottom:5px}
.sl .items{display:flex;flex-direction:column;gap:3px;max-height:236px;overflow:auto}
.wrap.p1 .sl{flex:1 1 190px}
.wrap.p1 .sl .items{flex-direction:row;flex-wrap:wrap;max-height:84px}
.sl .items button{text-align:left;border:0;border-radius:4px;padding:4px 8px;background:var(--on);color:var(--text);cursor:pointer;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:none;max-width:100%}
.sl .items button:hover{background:var(--hover)}
.sl .items button.off{background:var(--off);color:var(--muted)}
.sl .items button.none{opacity:.45}
.cards{display:grid;gap:var(--gap);grid-template-columns:repeat(var(--per),minmax(0,1fr));margin-bottom:var(--gap)}
.cards:empty{display:none}
.kpi{border-radius:var(--r);padding:11px 14px 10px;position:relative;box-shadow:var(--sh);overflow:hidden;min-width:0}
.kpi b{display:block;font-size:11px;letter-spacing:.04em;text-transform:uppercase;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi .row{display:flex;align-items:center;gap:10px;min-height:36px}
.kpi .v{font-size:26px;font-weight:700;line-height:1.25;white-space:nowrap}
.kpi svg{flex:1 1 40px;height:32px;min-width:0;max-width:46%;margin-left:auto}
.kpi small{display:block;font-size:11px;min-height:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi .bar{height:4px;border-radius:2px;margin-top:5px;overflow:hidden}
.kpi .bar i{display:block;height:100%;border-radius:2px}
.kpi .tl{position:absolute;left:14px;top:5px;width:34px;height:3px;border-radius:2px}
.sq .kpi .tl{left:0;top:0;width:100%;border-radius:0}
.kpi .sb{position:absolute;left:5px;top:10px;bottom:10px;width:4px;border-radius:2px}
.sq .kpi .sb{left:0;top:0;bottom:0;border-radius:0}
.kpi.side{padding-left:18px}
.charts{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:var(--gap)}
.ch{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:10px 12px;display:flex;flex-direction:column;min-width:0}
.ch .hd{display:flex;align-items:flex-start;gap:8px;min-height:24px}
.ch h2{margin:0;font-size:14px;line-height:1.3;font-weight:700;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ch .hd select,.ch .hd button{border:1px solid transparent;background:transparent;color:var(--muted);font-size:11px;border-radius:5px;padding:1px 3px;cursor:pointer;opacity:0;transition:opacity .15s}
.ch .hd select{max-width:118px}
.ch:hover .hd select,.ch:hover .hd button,.ch .hd select:focus,.ch .hd button:focus{opacity:1;border-color:var(--edge)}
.ch .hd option{color:#0f172a;background:#fff}
.ch .mid{flex:1;min-height:0;display:flex}
.ch .body{flex:1;min-width:0;min-height:40px;position:relative;overflow:hidden}
.ch .body.tbl{overflow:auto}
.ch svg{display:block;max-width:100%;height:auto}
.ch svg [data-v]{cursor:pointer}
.ch .empty{color:var(--muted);display:flex;align-items:center;justify-content:center;height:100%}
.legend{display:flex;flex-wrap:wrap;gap:3px 12px;font-size:11px;color:var(--muted);max-height:36px;overflow:hidden;flex:none}
.legend.top{padding-bottom:4px}
.legend.bottom{padding-top:4px}
.legend.side{flex-direction:column;flex-wrap:nowrap;max-height:none;justify-content:center;padding-left:10px;width:36%;max-width:180px}
.legend span{display:inline-flex;align-items:center;gap:5px;min-width:0}
.legend em{font-style:normal;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.legend i{width:9px;height:9px;border-radius:2px;flex:none}
table.t{width:100%;border-collapse:collapse;font-size:12px}
table.t th{text-align:left;font-weight:600;color:var(--muted);border-bottom:1px solid var(--grid);padding:4px 6px;position:sticky;top:0;background:var(--card);white-space:nowrap}
table.t td{padding:4px 6px;border-bottom:1px solid var(--grid);white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis}
table.t .n{text-align:right;font-variant-numeric:tabular-nums}
table.t td.b{width:34%}
table.t td.b i{display:block;height:10px;border-radius:2px;min-width:1px}
table.t tr[data-v]{cursor:pointer}
table.t tr.dim td{opacity:.4}
.ins{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:12px 16px;margin-top:var(--gap)}
.ins:empty{display:none}
.ins h3{margin:0 0 6px;font-size:11px;letter-spacing:.06em;color:var(--muted)}
.ins ul{margin:0;padding-left:18px}
.ins li{margin:3px 0}
footer{max-width:1240px;margin:0 auto;padding:10px 20px 18px;color:var(--muted);font-size:11px}
#tip{position:fixed;z-index:9;pointer-events:none;background:#0f172a;color:#fff;font-size:12px;line-height:1.35;padding:6px 9px;border-radius:6px;white-space:pre;box-shadow:0 4px 14px rgba(0,0,0,.25);display:none;max-width:340px;overflow:hidden;text-overflow:ellipsis}
@media(hover:none){.ch .hd select,.ch .hd button{opacity:1;border-color:var(--edge)}}
@media(max-width:980px){.wrap,.wrap.p2{grid-template-columns:minmax(0,1fr)}.wrap.p2 aside{order:0}aside{position:static;max-height:none}aside summary{display:block;cursor:pointer;font-weight:700;padding:8px 12px;background:var(--card);border-radius:var(--r);box-shadow:var(--sh)}aside details[open] summary{margin-bottom:var(--gap)}#filters{display:flex;flex-wrap:wrap}.sl{flex:1 1 180px}.sl .items{flex-direction:row;flex-wrap:wrap;max-height:96px}.ch{grid-column:span 12!important}.cards{grid-template-columns:repeat(2,minmax(0,1fr))}}
@media(min-width:981px) and (max-width:1180px){.ch.narrow{grid-column:span 6!important}}
@media(max-width:520px){.cards{grid-template-columns:minmax(0,1fr)}header{padding:12px 14px}.wrap{padding:10px 12px 6px}.chips{padding:8px 12px 0}}
@media print{.tools,.chips .all,.chip button,.ch .hd select,.ch .hd button,.sl h3 button,.sl input,aside summary{display:none!important}aside{position:static;max-height:none}.sl .items{max-height:none}.ch,.kpi,.sl,.ins{box-shadow:none;break-inside:avoid}.ch,.sl,.ins{border:1px solid #d0d7e2}}
</style>
</head>
<body>
<header id=""band""><div class=""hd""><h1 id=""title"">__TITLE__</h1><p id=""sub""></p></div><img id=""logo"" alt="""" hidden><div class=""tools"" id=""tools""></div></header>
<div class=""chips"" id=""chips""></div>
<div class=""wrap"" id=""wrap""><aside id=""side""><details id=""fbox"" open><summary>Filters</summary><div id=""filters""></div></details></aside><main><section class=""cards"" id=""cards""></section><section class=""charts"" id=""charts""></section><section class=""ins"" id=""insights""></section></main></div>
<footer id=""foot""></footer>
<div id=""tip""></div>
<script>
var D=/*DATA*/null;
(function(){
'use strict';
function $(id){return document.getElementById(id);}
if(!D||!D.cols){$('sub').textContent='This file has no data.';return;}

// ================= the data =================
var N=D.n,UNIT=D.unit,FY=D.fy,L=D.look,CH=D.charts,CA=D.cards;
var MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],DAY=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
var NODATE='(no date)',BLANK='(blank)',MAXS=8;
var PART={'@month':'Month','@quarter':'Quarter','@fy':'Fiscal Year','@year':'Year','@weekday':'Weekday','@day':'Day'};
var KINDS=['Column','Bar','Line','Area','Pie','Doughnut','Stacked column','Stacked bar','100% stacked','Combo','Pareto','Table'];
var WORDS=['Total','Average','Count of','Largest','Smallest','Different'];
var COL={};
D.cols.forEach(function(c){
 if(c.code){c.code=Int32Array.from(c.code);c.blank=c.items.indexOf(BLANK);}
 if(c.num){var a=new Float64Array(c.num.length);for(var i=0;i<a.length;i++)a[i]=c.num[i]===null?NaN:c.num[i];c.num=a;}
 COL[c.name.toLowerCase()]=c;
});
var YMD=D.ymd?Int32Array.from(D.ymd):null;
function col(n){return n?(COL[String(n).toLowerCase()]||null):null;}
function isPart(f){return !!f&&f.length>1&&f.charAt(0)==='@';}
function isStacked(k){return k===6||k===7||k===8;}
function isRound(k){return k===4||k===5;}
function isLine(k){return k===2||k===3;}
function canShowAs(k){return k>=0&&k<=3;}
function pad(n,w){var s=String(n);while(s.length<w)s='0'+s;return s;}

// ---- names of the groups of a date: these sort in the order of time ----
function fiscal(y,m){var f=m<FY?y-1:y;return FY===1?'FY '+f:'FY '+f+'-'+pad((f+1)%100,2);}
function partLabel(part,x){
 if(x<=0)return NODATE;
 var y=Math.floor(x/10000),m=Math.floor(x/100)%100,d=x%100;
 if(part==='@month')return pad(y,4)+'-'+pad(m,2)+' ('+MON[m-1]+')';
 if(part==='@quarter'){var q=Math.floor(((m-FY+12)%12)/3)+1;return FY===1?y+' Q'+q:fiscal(y,m)+' Q'+q;}
 if(part==='@fy')return fiscal(y,m);
 if(part==='@year')return pad(y,4);
 if(part==='@weekday'){var w=(new Date(y,m-1,d).getDay()+6)%7;return (w+1)+' '+DAY[w];}
 return pad(y,4)+'-'+pad(m,2)+'-'+pad(d,2);
}
// the short name shown on an axis: ""2026-04 (Apr)"" is shown as ""Apr 26""
function shortLabel(part,s){
 if(!isPart(part)||s===NODATE)return s;
 if(part==='@month')return s.length===13&&s.charAt(4)==='-'?s.substr(9,3)+' '+s.substr(2,2):s;
 if(part==='@quarter'){var q=s.lastIndexOf(' Q');if(q<0||q+3!==s.length)return s;var h=s.substring(0,q),t=s.substring(q+1);return h.indexOf('FY ')===0&&h.length===10?t+' '+h.substring(5):t+' '+h;}
 if(part==='@weekday')return s.length===5&&s.charAt(1)===' '?s.substring(2):s;
 if(part==='@day'){var m=parseInt(s.substr(5,2),10);return s.length===10&&m>=1&&m<=12?s.substr(8,2)+' '+MON[m-1]+' '+s.substr(2,2):s;}
 return s;
}
// one group number for every row, and the names of the groups
var KC={};
function keysOf(f){
 if(!f)return null;
 if(Object.prototype.hasOwnProperty.call(KC,f))return KC[f];
 var r=null,i;
 if(isPart(f)){
  if(YMD&&PART[f]){
   var key=new Int32Array(N),names=[],by=new Map();
   for(i=0;i<N;i++){
    var x=YMD[i],num;
    if(x<=0)num=0;
    else if(f==='@day')num=x;
    else if(f==='@weekday')num=1+(new Date(Math.floor(x/10000),Math.floor(x/100)%100-1,x%100).getDay()+6)%7;
    else if(f==='@year')num=Math.floor(x/10000);
    else if(f==='@month')num=Math.floor(x/100);
    else{var y=Math.floor(x/10000),m=Math.floor(x/100)%100,fy=m<FY?y-1:y;num=f==='@fy'?fy:fy*10+Math.floor(((m-FY+12)%12)/3)+1;}
    var id=by.get(num);
    if(id===undefined){id=names.length;by.set(num,id);names.push(partLabel(f,x));}
    key[i]=id;
   }
   r={key:key,labels:names};
  }
 }else{var c=col(f);if(c&&c.code)r={key:c.code,labels:c.items};}
 KC[f]=r;
 return r;
}

// ================= adding up (the same rules as the dashboard in Excel) =================
// key[row] = group number, or -1 to leave the row out. An empty column name counts rows.
function reduce(key,groups,name,agg){
 var res=new Float64Array(groups),c=col(name),i,g;
 if(groups===0)return res;
 if(!c){for(i=0;i<N;i++){g=key[i];if(g>=0)res[g]++;}return res;}
 if(agg===5){
  var seen=new Set();
  if(c.code){for(i=0;i<N;i++){g=key[i];if(g<0||c.code[i]===c.blank)continue;var k=g*20001+c.code[i];if(!seen.has(k)){seen.add(k);res[g]++;}}}
  else if(c.num){var ids=new Map();for(i=0;i<N;i++){g=key[i];var x=c.num[i];if(g<0||x!==x)continue;var id=ids.get(x);if(id===undefined){id=ids.size;ids.set(x,id);}var k2=g*2000000000+id;if(!seen.has(k2)){seen.add(k2);res[g]++;}}}
  return res;
 }
 if(agg===2){
  for(i=0;i<N;i++){
   g=key[i];if(g<0)continue;
   var f;
   if(c.num)f=c.num[i]===c.num[i];else if(c.code)f=c.code[i]!==c.blank;else if(c.filled)f=c.filled.charCodeAt(i)===49;else f=false;
   if(f)res[g]++;
  }
  return res;
 }
 if(!c.num)return res;
 var sum=new Float64Array(groups),cnt=new Int32Array(groups),mx=new Float64Array(groups),mn=new Float64Array(groups);
 for(i=0;i<N;i++){
  g=key[i];if(g<0)continue;
  var v=c.num[i];if(v!==v)continue;
  if(cnt[g]===0){mx[g]=v;mn[g]=v;}else{if(v>mx[g])mx[g]=v;if(v<mn[g])mn[g]=v;}
  sum[g]+=v;cnt[g]++;
 }
 for(g=0;g<groups;g++){
  if(agg===1)res[g]=cnt[g]>0?sum[g]/cnt[g]:0;
  else if(agg===3)res[g]=cnt[g]>0?mx[g]:0;
  else if(agg===4)res[g]=cnt[g]>0?mn[g]:0;
  else res[g]=sum[g];
 }
 return res;
}
function scalar(name,agg,mask){
 var key=new Int32Array(N);
 if(mask)for(var i=0;i<N;i++)if(!mask[i])key[i]=-1;
 return reduce(key,1,name,agg)[0];
}
function autoTop(kind,groups){
 if(isRound(kind))return groups>7?6:0;
 if(kind===1)return groups>12?10:0;
 if(kind===10)return groups>15?15:0;
 if(kind===11)return groups>10?10:0;
 if(isLine(kind)||kind===9)return groups>24?20:0;
 return groups>12?10:0;
}
function usesDate(c){
 if(!YMD)return false;
 if(isPart(c.dim))return true;
 return isStacked(c.kind)&&isPart(c.split)&&c.split!==c.dim;
}
function freeLabel(labels,want){
 var clash=true;
 while(clash){clash=false;for(var i=0;i<labels.length;i++)if(labels[i].toUpperCase()===want.toUpperCase()){clash=true;break;}if(clash)want+=' (rest)';}
 return want;
}
// by name: capitals and small letters are the same, as in a pivot table
function cmpName(a,b){var x=a.toUpperCase(),y=b.toUpperCase();if(x<y)return -1;if(x>y)return 1;return a<b?-1:(a>b?1:0);}
// mode 1 largest first, 2 smallest first, otherwise by name; equal ones keep their order
function sortIds(ids,val,labels,mode){
 var a=ids.map(function(g,i){return{g:g,i:i};});
 a.sort(function(x,y){
  var r=mode===1?val[y.g]-val[x.g]:(mode===2?val[x.g]-val[y.g]:cmpName(labels[x.g],labels[y.g]));
  return r<0?-1:(r>0?1:x.i-y.i);
 });
 return a.map(function(o){return o.g;});
}
function sortByValue(ids,val,labels,mode){return sortIds(sortIds(ids,val,labels,3),val,labels,mode);}

// what one chart shows
function group(c,mask){
 var res={cats:[],S:[],time:false,allCats:0,topN:0,others:false,othersName:'Others',sortMode:0,keep:[],splitKeep:[],splitOthers:false,splitOthersName:'Others',maxAbs:0};
 var K=keysOf(c.dim);
 if(!K)return res;
 var labels=K.labels,ng=labels.length,i,g;
 res.time=isPart(c.dim);
 var dated=usesDate(c),key=new Int32Array(N);
 for(i=0;i<N;i++)key[i]=((mask&&!mask[i])||(dated&&YMD[i]<=0))?-1:K.key[i];
 var rows=reduce(key,ng,'',2),val=reduce(key,ng,c.col,c.agg),ids=[];
 for(g=0;g<ng;g++)if(rows[g]>0)ids.push(g);
 res.allCats=ids.length;
 var mode=c.sort;
 if(mode===0){if(res.time||isLine(c.kind)||c.kind===9)mode=res.time?0:3;else mode=1;}
 if(c.kind===10)mode=1;
 res.sortMode=mode;
 var top=c.top;
 if(top===0)top=res.time?0:autoTop(c.kind,ids.length);
 if(res.time&&c.top<=0)top=0;
 if(c.showAs===2&&canShowAs(c.kind))top=0;
 var cut=top>0&&ids.length>top,keep=ids.slice();
 if(cut){
  keep=sortByValue(keep,val,labels,1);keep.length=top;res.topN=top;
  res.others=c.others===1||(c.others===0&&isRound(c.kind));
  if(c.kind===10)res.others=false;
 }
 keep=(mode===1||mode===2)?sortByValue(keep,val,labels,mode):sortIds(keep,val,labels,mode);
 var nk=keep.length,place=new Int32Array(ng).fill(-1);
 for(i=0;i<nk;i++){place[keep[i]]=i;res.keep.push(labels[keep[i]]);res.cats.push(labels[keep[i]]);}
 var total=nk;
 if(cut&&res.others){
  res.othersName=freeLabel(labels,'Others');
  for(g=0;g<ng;g++)if(place[g]<0&&rows[g]>0)place[g]=nk;
  res.cats.push(res.othersName);total=nk+1;
 }
 var k2=new Int32Array(N);
 for(i=0;i<N;i++)k2[i]=key[i]<0?-1:place[key[i]];
 if(isStacked(c.kind)&&keysOf(c.split)&&c.split!==c.dim)addSplit(c,res,k2,total);
 else{
  var s1={name:measure(c.col,c.agg),V:reduce(k2,total,c.col,c.agg),percent:false},all=0,run=0;
  res.S.push(s1);
  if(c.kind===9||(c.kind===11&&c.col2))res.S.push({name:measure(c.col2,c.agg2),V:reduce(k2,total,c.col2,c.agg2),percent:false});
  if(c.kind===10){
   var cum={name:'Cumulative share',V:new Float64Array(total),percent:true};
   for(i=0;i<total;i++)all+=s1.V[i];
   for(i=0;i<total;i++){run+=s1.V[i];cum.V[i]=all!==0?run/all:0;}
   res.S.push(cum);
  }else if(c.kind===11){
   var sh={name:'Share',V:new Float64Array(total),percent:true};
   for(i=0;i<total;i++)all+=s1.V[i];
   for(i=0;i<total;i++)sh.V[i]=all!==0?s1.V[i]/all:0;
   res.S.push(sh);
  }else if(c.showAs===1&&canShowAs(c.kind)){
   for(i=0;i<total;i++)all+=s1.V[i];
   for(i=0;i<total;i++)s1.V[i]=all!==0?s1.V[i]/all:0;
   s1.percent=true;
  }else if(c.showAs===2&&canShowAs(c.kind)){
   for(i=0;i<total;i++){run+=s1.V[i];s1.V[i]=run;}
  }
 }
 if(res.S.length)for(i=0;i<res.S[0].V.length;i++)res.maxAbs=Math.max(res.maxAbs,Math.abs(res.S[0].V[i]));
 return res;
}
// one series for each value of the colours column: the largest ones, the rest as Others
function addSplit(c,res,k2,total){
 var SK=keysOf(c.split),sl=SK.labels,sk=SK.key,ns=sl.length,i,g,j;
 var shown=new Int32Array(N);
 for(i=0;i<N;i++)shown[i]=k2[i]<0?-1:sk[i];
 var srows=reduce(shown,ns,'',2),sval=reduce(shown,ns,c.col,c.agg),sid=[];
 for(g=0;g<ns;g++)if(srows[g]>0)sid.push(g);
 var scut=sid.length>MAXS+1;
 if(scut){sid=sortByValue(sid,sval,sl,1);sid.length=MAXS;}
 sid=isPart(c.split)?sortIds(sid,sval,sl,0):sortByValue(sid,sval,sl,1);
 var nsk=sid.length,splace=new Int32Array(ns).fill(scut?nsk:-1);
 for(i=0;i<nsk;i++)splace[sid[i]]=i;
 var stotal=scut?nsk+1:nsk;
 res.splitOthers=scut;
 if(scut)res.splitOthersName=freeLabel(sl,'Others');
 var both=new Int32Array(N);
 for(i=0;i<N;i++){var a=k2[i],b=splace[sk[i]];both[i]=(a<0||b<0)?-1:a*stotal+b;}
 var cell=reduce(both,total*stotal,c.col,c.agg);
 for(j=0;j<stotal;j++){
  var se={name:j<nsk?sl[sid[j]]:res.splitOthersName,V:new Float64Array(total),percent:false};
  if(j<nsk)res.splitKeep.push(se.name);
  for(i=0;i<total;i++)se.V[i]=cell[i*stotal+j];
  res.S.push(se);
 }
 if(c.kind===8){
  for(i=0;i<total;i++){
   var all=0;
   for(j=0;j<stotal;j++)all+=res.S[j].V[i];
   for(j=0;j<stotal;j++)res.S[j].V[i]=all!==0?res.S[j].V[i]/all:0;
  }
  res.S.forEach(function(x){x.percent=true;});
 }
}
// the months of one number, oldest first
function months(name,agg,mask){
 var r=group({kind:2,dim:'@month',col:name,agg:agg,split:'',col2:'',agg2:0,top:0,others:0,sort:0,showAs:0},mask);
 if(r.cats.length&&r.cats[0]===NODATE&&r.S.length){r.cats.shift();r.S[0].V=r.S[0].V.slice(1);}
 return r;
}
// change of the latest month: mode 1 against the month before, 2 against the same month last year
function change(m,mode){
 var n=m.cats.length;
 if(n<2||!m.S.length)return NaN;
 var v=m.S[0].V,last=v[n-1],prev=NaN;
 if(mode===2){
  var cur=m.cats[n-1],y=parseInt(cur.substring(0,4),10);
  if(cur.length<7||isNaN(y))return NaN;
  var want=pad(y-1,4)+cur.substr(4,3);
  for(var i=0;i<n-1;i++)if(m.cats[i].indexOf(want)===0)prev=v[i];
 }else prev=v[n-2];
 if(prev!==prev||prev===0)return NaN;
 return (last-prev)/Math.abs(prev);
}

// ================= numbers as text =================
function fixed(v,d){
 if(!isFinite(v))return '0';
 var p=Math.pow(10,d),a=Math.round(Math.abs(v)*p*(1+1e-13))/p,s=a.toFixed(d);
 return (v<0&&parseFloat(s)!==0?'-':'')+s;
}
function whole(v){return v===Math.floor(v);}
function commas(s,indian){
 var neg=s.charAt(0)==='-';if(neg)s=s.substring(1);
 var dot=s.indexOf('.'),ip=dot<0?s:s.substring(0,dot),fp=dot<0?'':s.substring(dot);
 if(ip.length>3){
  var o=ip.substring(ip.length-3);ip=ip.substring(0,ip.length-3);
  var step=indian?2:3;
  while(ip.length>step){o=ip.substring(ip.length-step)+','+o;ip=ip.substring(0,ip.length-step);}
  ip=ip+','+o;
 }
 return (neg?'-':'')+ip+fp;
}
function indianGroup(v,dec){dec=Math.max(0,Math.min(6,dec));var s=commas(fixed(Math.abs(v),dec),true);return (v<0?'-':'')+s;}
function grouped(v,dec,indian){
 if(dec<0)dec=whole(v)?0:2;
 return indian?indianGroup(v,dec):commas(fixed(v,Math.min(dec,6)),false);
}
function shortNum(v,unit,dec){
 var a=Math.abs(v),d=dec<0?2:dec;
 if(unit===0){if(a>=1e7)return fixed(v/1e7,d)+' Cr';if(a>=1e5)return fixed(v/1e5,d)+' L';return grouped(v,dec,true);}
 if(unit===1){if(a>=1e9)return fixed(v/1e9,d)+' B';if(a>=1e6)return fixed(v/1e6,d)+' M';if(a>=1e3)return fixed(v/1e3,dec<0?1:dec)+' K';return grouped(v,dec,false);}
 return grouped(v,dec,false);
}
function trim0(s){if(s.indexOf('.')<0)return s;s=s.replace(/0+$/,'');return s.charAt(s.length-1)==='.'?s.substring(0,s.length-1):s;}
// short text for axes, labels and tables
function compact(v,unit,percent){
 if(percent)return fixed(v*100,Math.abs(v)<0.1&&v!==0?1:0)+'%';
 var a=Math.abs(v);
 if(unit===0){if(a>=1e7)return trim0(fixed(v/1e7,2))+' Cr';if(a>=1e5)return trim0(fixed(v/1e5,2))+' L';return indianGroup(v,a<100&&!whole(v)?1:0);}
 if(unit===1){if(a>=1e9)return trim0(fixed(v/1e9,2))+'B';if(a>=1e6)return trim0(fixed(v/1e6,2))+'M';if(a>=1e3)return trim0(fixed(v/1e3,1))+'K';}
 return commas(fixed(v,a<100&&!whole(v)?1:0),false);
}
function cardText(v,c){
 var core=c.mode===3?fixed(v*100,c.dec<0?1:c.dec)+'%':(c.mode===2?grouped(v,c.dec,c.indian):shortNum(v,c.mode,c.dec));
 return c.pre+core+c.suf;
}
function measure(name,agg){if(!name)return 'Rows';if(agg===5)return 'Different '+name;if(agg===2)return 'Count of '+name;return WORDS[agg]+' '+name;}
function shortM(name,agg){if(!name)return 'Rows';return agg===0?name:measure(name,agg);}
function chartTitle(c,r){
 if(c.title)return c.title;
 var m=shortM(c.col,c.agg),d=c.dimName,top=r&&r.topN>0?'Top '+r.topN+': ':'';
 if(c.kind===10)return 'Pareto: '+m+' by '+d;
 if(c.kind===9)return m+' and '+shortM(c.col2,c.agg2)+' by '+d;
 if(isStacked(c.kind)&&c.split)return top+m+' by '+d+' and '+c.splitName;
 var t=top+m+' by '+d;
 if(c.showAs===1)t+=' (% of total)';
 if(c.showAs===2)t+=' (running total)';
 return t;
}
// the line under the number of a card
function cardNote(c,mask){
 var v,n,ch;
 if(c.cmp>0&&YMD){
  ch=change(months(c.col,c.agg,mask),c.cmp);
  if(ch!==ch)return{text:'',sign:0};
  return{text:(ch>=0?'\u25B2 ':'\u25BC ')+fixed(Math.abs(ch)*100,1)+'% '+(c.cmp===2?'vs same month last year':'vs previous month'),sign:ch>=0?1:-1};
 }
 if(c.target>0){v=scalar(c.col,c.agg,mask);return{text:fixed(v/c.target*100,0)+'% of target '+shortNum(c.target,UNIT,-1),sign:0};}
 if(c.agg===0&&c.col){v=scalar(c.col,c.agg,mask);n=scalar('',2,mask);return{text:'avg per row: '+grouped(n>0?v/n:0,2,UNIT===0),sign:0};}
 if(!c.col)return{text:'rows after filters',sign:0};
 return{text:['','average of all rows','filled cells','largest value','smallest value','different values'][c.agg]||'',sign:0};
}
// a few plain sentences about the rows that are shown
function insights(mask){
 var res=[],name=D.ins.col,agg=name?0:2,i,dim=D.ins.dim,mName=name||'rows';
 if(dim&&keysOf(dim)){
  var r=group({kind:0,dim:dim,col:name,agg:agg,split:'',col2:'',agg2:0,top:-1,others:0,sort:1,showAs:0},mask),v=r.S.length?r.S[0].V:[],tot=0,n=v.length;
  for(i=0;i<n;i++)tot+=v[i];
  if(n>1&&tot>0){
   res.push(r.cats[0]+' leads '+dim+' with '+shortNum(v[0],UNIT,-1)+' ('+fixed(v[0]/tot*100,0)+'% of '+mName+').');
   if(n>=5){
    res.push('The top 3 '+dim+' values together make '+fixed((v[0]+v[1]+v[2])/tot*100,0)+'% of '+mName+'.');
    var cum=0,cnt=0;
    for(i=0;i<n;i++){cum+=v[i];cnt++;if(cum>=0.8*tot)break;}
    res.push(cnt+' of '+n+' '+dim+' values give 80% of '+mName+'.');
   }
   res.push('Smallest '+dim+': '+r.cats[n-1]+' ('+shortNum(v[n-1],UNIT,-1)+').');
  }
 }
 if(YMD){
  var m=months(name,agg,mask),k=m.cats.length;
  if(k>1){
   var w=m.S[0].V,best=0;
   for(i=1;i<k;i++)if(w[i]>w[best])best=i;
   res.push('Best month: '+m.cats[best]+' with '+shortNum(w[best],UNIT,-1)+'.');
   if(w[k-2]!==0){var ch=(w[k-1]-w[k-2])/Math.abs(w[k-2]);res.push('Latest month '+m.cats[k-1]+' is '+fixed(Math.abs(ch)*100,0)+'%'+(ch>=0?' up on ':' down on ')+m.cats[k-2]+'.');}
  }
 }
 return res;
}

// ================= colours =================
var TH=D.own;
function rgb(h){return[parseInt(h.substr(1,2),16),parseInt(h.substr(3,2),16),parseInt(h.substr(5,2),16)];}
function mix(a,b,t){var x=rgb(a),y=rgb(b),s='#';t=Math.max(0,Math.min(1,t));for(var i=0;i<3;i++)s+=pad(Math.round(x[i]+(y[i]-x[i])*t).toString(16),2);return s;}
function lum(h){var c=rgb(h);return(0.299*c[0]+0.587*c[1]+0.114*c[2])/255;}
function ink(bg){return lum(bg)>0.62?'#0f172a':'#ffffff';}
function readable(want,bg){return Math.abs(lum(want)-lum(bg))<0.2?TH.text:want;}
function pickColor(code,pos){if(typeof code==='string')return code;if(code>=0&&code<6)return TH.a[code];return TH.a[((pos%6)+6)%6];}
function pointColor(j){var c=TH.a[((j%6)+6)%6],r=Math.floor(j/6);if(r===1)c=mix(c,TH.card,0.4);else if(r>=2)c=mix(c,TH.text,0.35);return c;}
function othersColor(){return mix(TH.muted,TH.card,0.45);}
function seriesColor(c,idx,j){
 if(c.kind===9||c.kind===10){var first=pickColor(c.color,idx);if(j===0)return first;return TH.a[2]===first?TH.a[4]:TH.a[2];}
 return pointColor(j);
}
// the colours of one card, from the card style
function cardLook(style,accent){
 var k={bg:TH.card,bg2:TH.card,label:TH.muted,value:accent,note:TH.muted,line:accent,kind:0,grad:false,plain:true},i;
 if(style===0||style===5){
  i=ink(accent);
  return{bg:accent,bg2:style===5?mix(accent,i==='#ffffff'?'#000000':'#ffffff',0.38):accent,label:mix(accent,i,0.88),value:i,note:mix(accent,i,0.82),line:accent,kind:0,grad:style===5,plain:false};
 }
 k.value=readable(accent,TH.card);
 if(style===1)k.kind=1;
 else if(style===2)k.kind=2;
 else if(style===3){k.kind=3;k.value=TH.text;}
 else{
  k.bg=k.bg2=mix(TH.card,accent,TH.d?0.24:0.12);
  k.value=readable(TH.d?mix(accent,'#ffffff',0.25):mix(accent,'#000000',0.2),k.bg);
  k.label=mix(TH.muted,k.value,0.35);
 }
 return k;
}

// ================= what the viewer has chosen =================
var SEL=new Map();                        // field -> Set of the names that are switched on
var st={kind:CH.map(function(c){return c.kind;}),cards:L.cards,round:L.round,theme:D.theme};
function chartNow(i){var c=CH[i];if(st.kind[i]===c.kind)return c;var o={};for(var k in c)o[k]=c[k];o.kind=st.kind[i];return o;}
// 1 for every row that passes all filters; the filter on 'skip' is left out. null = every row.
function maskOf(skip){
 var m=null;
 SEL.forEach(function(set,f){
  if(f===skip||!set.size)return;
  var K=keysOf(f);if(!K)return;
  var ok=new Uint8Array(K.labels.length),i;
  for(i=0;i<ok.length;i++)if(set.has(K.labels[i]))ok[i]=1;
  if(!m)m=new Uint8Array(N).fill(1);
  for(i=0;i<N;i++)if(!ok[K.key[i]])m[i]=0;
 });
 return m;
}
function toggle(f,v){
 var set=SEL.get(f);
 if(!set){set=new Set();SEL.set(f,set);}
 if(set.has(v))set.delete(v);else set.add(v);
 if(!set.size)SEL.delete(f);
}
function fieldName(f){return isPart(f)?PART[f]:f;}

// ================= drawing =================
function esc(s){return String(s).replace(/[&<>""']/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','""':'&quot;',""'"":'&#39;'}[c];});}
var mctx=document.createElement('canvas').getContext('2d'),FONT='';
function tw(s,px,bold){mctx.font=(bold?'700 ':'400 ')+px+'px '+FONT;return mctx.measureText(s).width;}
function maxW(list,px){var w=0;for(var i=0;i<list.length;i++)w=Math.max(w,tw(list[i],px));return w;}
function fitText(s,w,px){if(tw(s,px)<=w)return s;while(s.length>1&&tw(s+'\u2026',px)>w)s=s.substring(0,s.length-1);return s+'\u2026';}
function f1(x){return Math.round(x*10)/10;}
function niceMax(v){if(!(v>0)||!isFinite(v))return 1;var p=Math.pow(10,Math.floor(Math.log(v)/Math.LN10+1e-12)),f=v/p;return(f<=1?1:(f<=2?2:(f<=2.5?2.5:(f<=5?5:10))))*p;}
// a value axis with round numbers that holds lo to hi; zero is always on it
function axis(lo,hi,percent){
 if(!(lo<=0))lo=0;if(!(hi>=0))hi=0;
 if(!isFinite(lo))lo=0;if(!isFinite(hi))hi=0;
 if(percent&&lo===0&&hi<=1.0000001&&hi>0.5)return{min:0,max:1,step:0.25};
 var span=hi-lo;
 if(!(span>0))return{min:0,max:1,step:0.25};
 var step=niceMax(span/5),a={step:step,max:Math.ceil(hi/step-1e-9)*step,min:-Math.ceil(-lo/step-1e-9)*step};
 if(a.max===a.min)a.max=a.min+step;
 return a;
}
function ticks(a){var n=Math.round((a.max-a.min)/a.step),t=[];for(var i=0;i<=n;i++)t.push(a.min+a.step*i);return t;}
function where(a,v){return a.max>a.min?(v-a.min)/(a.max-a.min):0;}
function axisFor(S,from,to,stacked,n){
 var mx=0,mn=0,i,j,v;
 for(i=0;i<n;i++){
  if(stacked){var up=0,dn=0;for(j=from;j<=to;j++){v=S[j].V[i];if(v>=0)up+=v;else dn+=v;}mx=Math.max(mx,up);mn=Math.min(mn,dn);}
  else for(j=from;j<=to;j++){v=S[j].V[i];mx=Math.max(mx,v);mn=Math.min(mn,v);}
 }
 return axis(mn,mx,S[from].percent);
}
function labelMode(c){if(c.labels>0)return c.labels===3&&!isRound(c.kind)?2:c.labels;return isRound(c.kind)?3:1;}
function legendPos(c,series){if(c.legend>0)return c.legend;if(c.kind===11)return 1;if(isRound(c.kind))return 2;return series>1?3:1;}
function svgText(x,y,s,anchor,px,fill,extra){return '<text x=""'+f1(x)+'"" y=""'+f1(y)+'"" text-anchor=""'+anchor+'"" font-size=""'+px+'"" style=""fill:'+fill+'""'+(extra||'')+'>'+esc(s)+'</text>';}
function tipAttr(lines){return ' data-tip=""'+esc(lines.join('\n'))+'""';}
function pickAttr(c,a,k){
 // a click on a group filters by it; the group ""Others"" cannot be clicked
 if(a.others&&k===a.cats.length-1)return '';
 return ' data-f=""'+esc(c.dim)+'"" data-v=""'+esc(a.cats[k])+'""';
}
function faded(c,a,k){var set=SEL.get(c.dim);return set&&set.size&&!set.has(a.cats[k])?' opacity=""0.3""':'';}
function shareOf(a,k){var t=0,V=a.S[0].V;for(var i=0;i<V.length;i++)t+=Math.abs(V[i]);return t>0?Math.abs(V[k])/t:0;}
function tipLines(c,a,k,j){
 var lines=[shortLabel(c.dim,a.cats[k])];
 if(j===undefined){for(var q=0;q<a.S.length;q++)lines.push(a.S[q].name+': '+compact(a.S[q].V[k],UNIT,a.S[q].percent));}
 else lines.push(a.S[j].name+': '+compact(a.S[j].V[k],UNIT,a.S[j].percent));
 return lines;
}

// columns, stacked columns, lines, areas, combo and Pareto charts
function drawCols(c,a,colors,W,H){
 var n=a.cats.length,S=a.S,kind=c.kind,line=isLine(kind),stacked=isStacked(kind),combo=(kind===9||kind===10)&&S.length>1;
 var bars=combo?1:S.length,pct=S[0].percent,ax=axisFor(S,0,bars-1,stacked,n),ax2=combo?axisFor(S,1,1,false,n):null,fs=11,i,j,k,s='';
 var tk=ticks(ax),tl=tk.map(function(v){return compact(v,UNIT,pct);}),left=maxW(tl,fs)+10,right=8,tk2,tl2;
 if(ax2){tk2=ticks(ax2);tl2=tk2.map(function(v){return compact(v,UNIT,S[1].percent);});right=maxW(tl2,fs)+10;}
 var mode=labelMode(c),top=mode===2&&!stacked?16:8,bottom=20,pw=W-left-right,ph=H-top-bottom;
 if(pw<30||ph<30)return '';
 var slot=pw/n,X=function(q){return left+slot*(q+0.5);},Y=function(v){return top+ph*(1-where(ax,v));},Y2=function(v){return top+ph*(1-where(ax2,v));},zero=Y(0);
 for(i=0;i<tk.length;i++){var gy=Math.round(Y(tk[i]))+0.5;s+='<line x1=""'+left+'"" x2=""'+f1(W-right)+'"" y1=""'+gy+'"" y2=""'+gy+'"" style=""stroke:var(--grid)""/>'+svgText(left-6,gy+3.5,tl[i],'end',fs,'var(--muted)');}
 if(ax2)for(i=0;i<tk2.length;i++)s+=svgText(W-right+6,Math.round(Y2(tk2[i]))+4,tl2[i],'start',fs,'var(--muted)');
 var names=a.cats.map(function(x){return shortLabel(c.dim,x);}),wide=Math.min(maxW(names,fs),96),every=Math.max(1,Math.ceil((wide+8)/slot));
 for(k=0;k<n;k+=every)s+=svgText(X(k),H-5,fitText(names[k],Math.max(slot*every-4,26),fs),'middle',fs,'var(--muted)',pickAttr(c,a,k)+tipAttr(tipLines(c,a,k)));
 var V=S[0].V,labelEvery=1;
 if(mode===2&&!stacked){var lw=0;for(k=0;k<Math.min(n,24);k++)lw=Math.max(lw,tw(compact(V[k],UNIT,pct),10));labelEvery=Math.max(1,Math.ceil((lw+12)/slot));}
 if(line){
  var pts=[],d='';
  for(k=0;k<n;k++)pts.push([X(k),Y(V[k])]);
  if(c.smooth&&n>2){
   d='M'+f1(pts[0][0])+' '+f1(pts[0][1]);
   for(k=0;k<n-1;k++){
    var p0=pts[Math.max(k-1,0)],p1=pts[k],p2=pts[k+1],p3=pts[Math.min(k+2,n-1)],t=0.4/3;
    d+='C'+f1(p1[0]+(p2[0]-p0[0])*t)+' '+f1(p1[1]+(p2[1]-p0[1])*t)+' '+f1(p2[0]-(p3[0]-p1[0])*t)+' '+f1(p2[1]-(p3[1]-p1[1])*t)+' '+f1(p2[0])+' '+f1(p2[1]);
   }
  }else d='M'+pts.map(function(p){return f1(p[0])+' '+f1(p[1]);}).join('L');
  if(kind===3&&n>1)s+='<path d=""'+d+'L'+f1(pts[n-1][0])+' '+f1(zero)+'L'+f1(pts[0][0])+' '+f1(zero)+'Z"" fill=""'+colors[0]+'"" fill-opacity=""0.27""/>';
  if(n>1)s+='<path d=""'+d+'"" fill=""none"" stroke=""'+colors[0]+'"" stroke-width=""2.25"" stroke-linejoin=""round"" stroke-linecap=""round""/>';
  if(c.trend>0&&n>2)s+=trendLine(V,c.trend,X,Y);
  for(k=0;k<n;k++){
   var show=(kind!==3&&slot>7)||n===1;
   s+='<circle cx=""'+f1(pts[k][0])+'"" cy=""'+f1(pts[k][1])+'"" r=""'+(show?3:Math.max(3,Math.min(slot/2,9)))+'""'+(show?' fill=""'+colors[0]+'""':' fill=""transparent""')+faded(c,a,k)+pickAttr(c,a,k)+tipAttr(tipLines(c,a,k))+'/>';
   if(mode===2&&k%labelEvery===0)s+=svgText(pts[k][0],pts[k][1]-7,compact(V[k],UNIT,pct),'middle',10,'var(--text)');
  }
 }else{
  var gw=slot*0.66,bw=stacked||combo?gw:gw/bars;
  for(k=0;k<n;k++){
   var x0=left+slot*k+(slot-gw)/2,up=0,dn=0;
   for(j=0;j<bars;j++){
    var v=S[j].V[k],from=0;
    if(stacked){from=v>=0?up:dn;if(v>=0)up+=v;else dn+=v;}
    var y1=Y(from+v),y0=Y(from),h=Math.abs(y1-y0),x=stacked||combo?x0:x0+bw*j,fill=bars>1?colors[j%colors.length]:colors[0];
    if(h>=0.5)s+='<rect x=""'+f1(x)+'"" y=""'+f1(Math.min(y0,y1))+'"" width=""'+f1(Math.max(1,bw-(stacked||combo?0:1)))+'"" height=""'+f1(h)+'"" fill=""'+fill+'""'+(stacked?'':' rx=""1.5""')+faded(c,a,k)+pickAttr(c,a,k)+tipAttr(tipLines(c,a,k,bars>1?j:undefined))+'/>';
    if(mode===2&&bars===1&&k%labelEvery===0)s+=svgText(x+bw/2,(v>=0?Math.min(y0,y1)-4:Math.max(y0,y1)+11),compact(v,UNIT,pct),'middle',10,'var(--text)');
   }
  }
  if(combo){
   var c2=colors[1%colors.length],V2=S[1].V,d2='';
   for(k=0;k<n;k++)d2+=(k?'L':'M')+f1(X(k))+' '+f1(Y2(V2[k]));
   if(n>1)s+='<path d=""'+d2+'"" fill=""none"" stroke=""'+c2+'"" stroke-width=""2.25"" stroke-linejoin=""round""/>';
   for(k=0;k<n;k++)if(slot>7||n<2)s+='<circle cx=""'+f1(X(k))+'"" cy=""'+f1(Y2(V2[k]))+'"" r=""3"" fill=""'+c2+'""'+tipAttr(tipLines(c,a,k,1))+'/>';
  }
 }
 return s;
}
// a straight trend line, or the moving average of 3 points
function trendLine(V,kind,X,Y){
 var n=V.length,d='',i;
 if(kind===2){for(i=2;i<n;i++)d+=(d?'L':'M')+f1(X(i))+' '+f1(Y((V[i]+V[i-1]+V[i-2])/3));}
 else{
  var sx=0,sy=0,sxx=0,sxy=0;
  for(i=0;i<n;i++){sx+=i;sy+=V[i];sxx+=i*i;sxy+=i*V[i];}
  var den=n*sxx-sx*sx;if(den===0)return '';
  var b=(n*sxy-sx*sy)/den,a0=(sy-b*sx)/n;
  d='M'+f1(X(0))+' '+f1(Y(a0))+'L'+f1(X(n-1))+' '+f1(Y(a0+b*(n-1)));
 }
 return d?'<path d=""'+d+'"" fill=""none"" style=""stroke:var(--muted)"" stroke-width=""1.25"" stroke-dasharray=""5 4""/>':'';
}
// bars and stacked bars: the first group on top
function drawBars(c,a,colors,W,H){
 var n=a.cats.length,S=a.S,stacked=isStacked(c.kind),pct=S[0].percent,ax=axisFor(S,0,S.length-1,stacked,n),fs=11,mode=labelMode(c),k,j,s='';
 var names=a.cats.map(function(x){return shortLabel(c.dim,x);}),gutter=Math.min(maxW(names,fs)+10,W*0.38);
 var tk=ticks(ax),tl=tk.map(function(v){return compact(v,UNIT,pct);}),numW=maxW(tl,fs),right=mode===2&&!stacked?numW+12:numW/2+4,top=4,bottom=20,pw=W-gutter-right,ph=H-top-bottom;
 if(pw<30||ph<20)return '';
 var X=function(v){return gutter+pw*where(ax,v);},slot=ph/n,bh=Math.max(1,Math.min(slot*0.66,34)),skip=Math.max(1,Math.ceil((numW+10)/(pw/Math.max(1,tk.length-1))));
 for(k=0;k<tk.length;k++){var gx=Math.round(X(tk[k]))+0.5;s+='<line x1=""'+gx+'"" x2=""'+gx+'"" y1=""'+top+'"" y2=""'+f1(top+ph)+'"" style=""stroke:var(--grid)""/>';if(k%skip===0)s+=svgText(gx,H-5,tl[k],'middle',fs,'var(--muted)');}
 var nameEvery=Math.max(1,Math.ceil(13/slot));
 for(k=0;k<n;k++){
  var yc=top+slot*(k+0.5),lo=0,hi=0;
  if(k%nameEvery===0)s+=svgText(gutter-6,yc+3.5,fitText(names[k],gutter-10,fs),'end',fs,'var(--muted)',pickAttr(c,a,k)+tipAttr(tipLines(c,a,k)));
  for(j=0;j<S.length;j++){
   var v=S[j].V[k],from=0;
   if(stacked){from=v>=0?hi:lo;if(v>=0)hi+=v;else lo+=v;}
   var x0=X(from),x1=X(from+v),w=Math.abs(x1-x0),y=stacked?yc-bh/2:yc-bh/2+(bh/S.length)*j,h=stacked?bh:bh/S.length;
   if(w>=0.5)s+='<rect x=""'+f1(Math.min(x0,x1))+'"" y=""'+f1(y)+'"" width=""'+f1(w)+'"" height=""'+f1(Math.max(1,h-(stacked||S.length===1?0:1)))+'"" fill=""'+(S.length>1?colors[j%colors.length]:colors[0])+'""'+(stacked?'':' rx=""1.5""')+faded(c,a,k)+pickAttr(c,a,k)+tipAttr(tipLines(c,a,k,S.length>1?j:undefined))+'/>';
   if(mode===2&&!stacked&&S.length===1&&slot>=11)s+=svgText(Math.max(x0,x1)+5,yc+3.5,compact(v,UNIT,pct),'start',10,'var(--text)');
  }
 }
 return s;
}
// pie and doughnut
function drawRound(c,a,colors,W,H){
 var V=a.S[0].V,n=V.length,tot=0,i,s='',d=Math.min(W,H)-6;
 for(i=0;i<n;i++)tot+=Math.abs(V[i]);
 if(d<24||tot<=0)return '';
 var cx=W/2,cy=H/2,r=d/2,hole=c.kind===5?r*0.58:0,ang=-Math.PI/2,mode=labelMode(c);
 var P=function(rad,t){return f1(cx+Math.cos(t)*rad)+' '+f1(cy+Math.sin(t)*rad);};
 for(i=0;i<n;i++){
  var sw=Math.abs(V[i])/tot*Math.PI*2;
  if(sw<=0)continue;
  var a0=ang,a1=ang+sw,big=sw>Math.PI?1:0,path;
  if(sw>=Math.PI*2-1e-9)path='M'+P(r,0)+'A'+r+' '+r+' 0 1 1 '+P(r,Math.PI)+'A'+r+' '+r+' 0 1 1 '+P(r,0)+(hole?'M'+P(hole,0)+'A'+hole+' '+hole+' 0 1 0 '+P(hole,Math.PI)+'A'+hole+' '+hole+' 0 1 0 '+P(hole,0):'')+'Z';
  else path='M'+P(r,a0)+'A'+f1(r)+' '+f1(r)+' 0 '+big+' 1 '+P(r,a1)+(hole?'L'+P(hole,a1)+'A'+f1(hole)+' '+f1(hole)+' 0 '+big+' 0 '+P(hole,a0):'L'+f1(cx)+' '+f1(cy))+'Z';
  s+='<path d=""'+path+'"" fill=""'+colors[i%colors.length]+'"" fill-rule=""evenodd"" style=""stroke:var(--card)"" stroke-width=""1.5""'+faded(c,a,i)+pickAttr(c,a,i)+tipAttr([shortLabel(c.dim,a.cats[i]),a.S[0].name+': '+compact(V[i],UNIT,false),'Share: '+fixed(Math.abs(V[i])/tot*100,1)+'%'])+'/>';
  if(mode!==1&&sw>=0.384){
   var mid=ang+sw/2,lr=c.kind===5?r*0.79:r*0.64;
   s+=svgText(cx+Math.cos(mid)*lr,cy+Math.sin(mid)*lr+3.5,mode===3?fixed(Math.abs(V[i])/tot*100,0)+'%':compact(V[i],UNIT,false),'middle',10.5,ink(colors[i%colors.length]),' font-weight=""700"" pointer-events=""none""');
  }
  ang=a1;
 }
 if(hole)s+=svgText(cx,cy+5,compact(tot,UNIT,false),'middle',Math.max(10,Math.min(16,hole*0.36)),'var(--text)',' font-weight=""700""');
 return s;
}
// a table with a bar behind the first number
function drawTable(c,a,idx){
 var S=a.S,n=a.cats.length,color=pickColor(c.color,idx),mx=0,k,j,set=SEL.get(c.dim);
 for(k=0;k<n;k++)mx=Math.max(mx,Math.abs(S[0].V[k]));
 var h='<table class=""t""><thead><tr><th>'+esc(c.dimName)+'</th><th></th>';
 for(j=0;j<S.length;j++)h+='<th class=""n"">'+esc(S[j].name)+'</th>';
 h+='</tr></thead><tbody>';
 for(k=0;k<n;k++){
  var others=a.others&&k===n-1;
  h+='<tr'+(others?'':' data-f=""'+esc(c.dim)+'"" data-v=""'+esc(a.cats[k])+'""')+(set&&set.size&&!set.has(a.cats[k])?' class=""dim""':'')+'><td>'+esc(shortLabel(c.dim,a.cats[k]))+'</td><td class=""b""><i style=""background:'+(others?othersColor():color)+';width:'+f1(mx>0?Math.abs(S[0].V[k])/mx*100:0)+'%""></i></td>';
  for(j=0;j<S.length;j++)h+='<td class=""n"">'+esc(compact(S[j].V[k],UNIT,S[j].percent))+'</td>';
  h+='</tr>';
 }
 return h+'</tbody></table>';
}
// the small trend line on a card
function spark(V,color){
 var n=V.length,mn=Infinity,mx=-Infinity,i,d='';
 if(n<2)return '';
 for(i=0;i<n;i++){mn=Math.min(mn,V[i]);mx=Math.max(mx,V[i]);}
 for(i=0;i<n;i++)d+=(i?'L':'M')+f1(100*i/(n-1))+' '+f1(mx>mn?28-(V[i]-mn)/(mx-mn)*24:16);
 return '<svg viewBox=""0 0 100 32"" preserveAspectRatio=""none"" aria-hidden=""true""><path d=""'+d+'"" fill=""none"" stroke=""'+color+'"" stroke-width=""1.8"" stroke-linejoin=""round"" vector-effect=""non-scaling-stroke""/></svg>';
}

// ================= the parts of the page =================
function drawCards(){
 var m=maskOf(null),h='';
 CA.forEach(function(c,i){
  var accent=pickColor(c.color,i),k=cardLook(st.cards,accent),v=scalar(c.col,c.agg,m),note=cardNote(c,m);
  var noteColor=k.plain&&note.sign?(note.sign>0?(TH.d?'#4ade80':'#16a34a'):(TH.d?'#f87171':'#dc2626')):k.note;
  var bg=k.grad?'linear-gradient(90deg,'+k.bg+','+k.bg2+')':k.bg;
  h+='<div class=""kpi'+(k.kind===2?' side':'')+'"" style=""background:'+bg+(k.kind===3?';box-shadow:inset 0 0 0 1.5px '+mix(TH.card,k.line,0.6):'')+'"">';
  if(k.kind===1)h+='<i class=""tl"" style=""background:'+k.line+'""></i>';
  if(k.kind===2)h+='<i class=""sb"" style=""background:'+k.line+'""></i>';
  h+='<b style=""color:'+k.label+'"">'+esc(c.label)+'</b><div class=""row""><span class=""v"" style=""color:'+k.value+'"">'+esc(cardText(v,c))+'</span>';
  if(c.spark&&YMD)h+=spark(months(c.col,c.agg,m).S[0].V,k.plain?k.line:k.value);
  h+='</div><small style=""color:'+noteColor+'"">'+esc(note.text)+'</small>';
  if(c.target>0&&!(c.cmp>0&&YMD))h+='<div class=""bar"" style=""background:'+mix(k.bg,k.value,0.18)+'""><i style=""background:'+k.value+';width:'+f1(Math.max(0,Math.min(1,v/c.target))*100)+'%""></i></div>';
  h+='</div>';
 });
 $('cards').innerHTML=h;
 // a long number is written smaller so that it fits its card
 var list=$('cards').querySelectorAll('.v');
 for(var i=0;i<list.length;i++){var e=list[i],size=26,room=e.parentNode.clientWidth-(e.nextSibling?46:0);while(size>13&&e.scrollWidth>room){size-=1;e.style.fontSize=size+'px';}}
}
var results=[];                           // what every chart shows now (also read by tests)
function drawCharts(){
 var host=$('charts'),h='';
 results=[];
 CH.forEach(function(c0,i){
  var c=chartNow(i),a=group(c,maskOf(c.dim));
  results.push(a);
  var round=isRound(c.kind),names=[],colors=[],j;
  if(round)for(j=0;j<a.cats.length;j++){names.push(shortLabel(c.dim,a.cats[j]));colors.push(a.others&&j===a.cats.length-1?othersColor():pointColor(j));}
  else if(a.S.length>1&&c.kind!==11)for(j=0;j<a.S.length;j++){names.push(isStacked(c.kind)?shortLabel(c.split,a.S[j].name):a.S[j].name);colors.push(a.splitOthers&&j===a.S.length-1?othersColor():seriesColor(c,i,j));}
  else{names.push(a.S.length?a.S[0].name:'');colors.push(pickColor(c.color,i));}
  var pos=legendPos(c,round?a.cats.length:a.S.length),leg='';
  if(pos!==1&&a.cats.length){
   leg='<div class=""legend '+(pos===2?'side':(pos===4?'top':'bottom'))+'"">';
   for(j=0;j<names.length;j++)leg+='<span><i style=""background:'+colors[j]+'""></i><em>'+esc(names[j])+'</em></span>';
   leg+='</div>';
  }
  var alts=kindsFor(c0),pick='';
  if(alts.length>1){
   pick='<select data-chart=""'+i+'"" title=""Kind of chart"" aria-label=""Kind of chart"">';
   alts.forEach(function(k){pick+='<option value=""'+k+'""'+(k===c.kind?' selected':'')+'>'+KINDS[k]+'</option>';});
   pick+='</select>';
  }
  h+='<div class=""ch'+(c0.w<=4?' narrow':'')+'"" style=""grid-column:span '+c0.w+';height:'+c0.h+'px"" data-i=""'+i+'""><div class=""hd""><h2 title=""'+esc(chartTitle(c,a))+'"">'+esc(chartTitle(c,a))+'</h2>'+pick+'<button type=""button"" data-csv=""'+i+'"" title=""Save the numbers of this chart as a CSV file"">CSV</button></div>'+
     (pos===4?leg:'')+'<div class=""mid""><div class=""body'+(c.kind===11?' tbl':'')+'""></div>'+(pos===2?leg:'')+'</div>'+(pos===3?leg:'')+'</div>';
  results[i].colors=colors;
 });
 host.innerHTML=h;
 var cards=host.children;
 for(var i=0;i<cards.length;i++){
  var c=chartNow(i),a=results[i],body=cards[i].querySelector('.body');
  if(!a.cats.length||!a.S.length){body.innerHTML='<div class=""empty"">Nothing to show</div>';continue;}
  if(c.kind===11){body.innerHTML=drawTable(c,a,i);continue;}
  var W=body.clientWidth,H=body.clientHeight;
  var inner=isRound(c.kind)?drawRound(c,a,a.colors,W,H):(c.kind===1||c.kind===7?drawBars(c,a,a.colors,W,H):drawCols(c,a,a.colors,W,H));
  body.innerHTML=inner?'<svg width=""'+W+'"" height=""'+H+'"" viewBox=""0 0 '+W+' '+H+'"" role=""img"" aria-label=""'+esc(chartTitle(c,a))+'"">'+inner+'</svg>':'<div class=""empty"">Not enough room</div>';
 }
}
// the kinds a viewer can switch a chart to: the ones that need no other columns
function kindsFor(c){
 if(isStacked(c.kind))return[6,7,8];
 if(c.kind===9)return[9];
 var k=[0,1,2,3,4,5,11];
 if(c.agg===0||c.agg===2)k.push(10);
 if(k.indexOf(c.kind)<0)k.push(c.kind);
 return k;
}
function drawFilters(){
 var h='';
 D.slicers.forEach(function(sl,si){
  var K=keysOf(sl.f);
  if(!K)return;
  var set=SEL.get(sl.f),rows=reduce(maskedKey(K,maskOf(sl.f)),K.labels.length,'',2);
  var order=K.labels.map(function(x,i){return i;}).sort(function(x,y){return cmpName(K.labels[x],K.labels[y]);});
  h+='<div class=""sl'+(set&&set.size?' has':'')+'"" data-f=""'+esc(sl.f)+'""><h3><span>'+esc(sl.name)+'</span><button type=""button"" data-clear=""'+esc(sl.f)+'"">clear</button></h3>';
  if(order.length>14)h+='<input type=""search"" placeholder=""Search"" aria-label=""Search '+esc(sl.name)+'"" data-find=""'+si+'"" value=""'+esc(finds[si]||'')+'"">';
  h+='<div class=""items"">';
  var q=(finds[si]||'').toLowerCase(),shown=0;
  order.forEach(function(g){
   var name=K.labels[g],label=shortLabel(sl.f,name);
   if(q&&label.toLowerCase().indexOf(q)<0)return;
   if(++shown>600)return;
   h+='<button type=""button"" data-f=""'+esc(sl.f)+'"" data-v=""'+esc(name)+'"" class=""'+(set&&set.size&&!set.has(name)?'off':'')+(rows[g]>0?'':' none')+'"" title=""'+esc(label)+'"">'+esc(label)+'</button>';
  });
  h+='</div></div>';
 });
 $('filters').innerHTML=h;
 var n=0;SEL.forEach(function(set){if(set.size)n++;});
 $('fbox').firstChild.textContent='Filters'+(n?' ('+n+' set)':'');
}
var finds={};
function maskedKey(K,m){if(!m)return K.key;var k=new Int32Array(N);for(var i=0;i<N;i++)k[i]=m[i]?K.key[i]:-1;return k;}
function drawChips(){
 var h='';
 SEL.forEach(function(set,f){
  if(!set.size)return;
  var names=[];set.forEach(function(v){names.push(shortLabel(f,v));});
  h+='<span class=""chip""><span><b>'+esc(fieldName(f))+':</b> '+esc(names.length>4?names.slice(0,4).join(', ')+' +'+(names.length-4):names.join(', '))+'</span><button type=""button"" data-clear=""'+esc(f)+'"" title=""Remove this filter"" aria-label=""Remove this filter"">\u00D7</button></span>';
 });
 if(h){var m=maskOf(null),n=0;if(m)for(var i=0;i<N;i++)n+=m[i];else n=N;h='<span>'+commas(String(n),UNIT===0)+' of '+commas(String(N),UNIT===0)+' rows</span>'+h+'<button type=""button"" class=""all"" data-clear=""*"">Clear all filters</button>';}
 $('chips').innerHTML=h;
}
function drawInsights(){
 var h='';
 if(L.insights){var list=insights(maskOf(null));if(list.length){h='<h3>KEY INSIGHTS</h3><ul>';list.forEach(function(x){h+='<li>'+esc(x)+'</li>';});h+='</ul>';}}
 $('insights').innerHTML=h;
}
function drawAll(){drawChips();drawFilters();drawCards();drawCharts();drawInsights();}

// ---- the look ----
function applyLook(){
 var r=document.documentElement.style,b=document.body,i;
 // the theme the design came with keeps its exact colours, also when they were changed by hand
 TH=(st.theme===D.theme||st.theme<0||st.theme>=D.themes.length)?D.own:D.themes[st.theme];
 ['page','band','bandText','card','text','muted','grid'].forEach(function(k){r.setProperty('--'+k,TH[k]);});
 r.setProperty('--on',mix(TH.card,TH.a[0],TH.d?0.42:0.2));
 r.setProperty('--off',mix(TH.card,TH.muted,0.08));
 r.setProperty('--hover',mix(TH.card,TH.a[0],TH.d?0.6:0.36));
 r.setProperty('--edge',mix(TH.card,TH.muted,0.3));
 r.setProperty('--r',st.round?'10px':'0px');
 r.setProperty('--sh',L.shadow?(TH.d?'0 1px 2px rgba(0,0,0,.5),0 3px 10px rgba(0,0,0,.35)':'0 1px 2px rgba(15,23,42,.10),0 2px 8px rgba(15,23,42,.06)'):'none');
 r.setProperty('--gap',[10,12,16][L.dens]+'px');
 r.setProperty('--per',String(Math.max(1,L.per||CA.length||1)));
 r.setProperty('--font',""'""+String(L.font).replace(/['\\]/g,'')+""'"");
 FONT=""'""+String(L.font).replace(/['\\]/g,'')+""','Segoe UI',system-ui,-apple-system,'Helvetica Neue',Arial,sans-serif"";
 b.className=(st.round?'':'sq ')+(L.shadow?'':'flat ')+(L.wide?'wide':'');
 $('band').className=lum(TH.band)>0.9&&lum(TH.page)>0.9?'line':'';
 $('wrap').className='wrap p'+(D.slicers.length?L.pos:3);
}
function drawTools(){
 var h='<select id=""pickTheme"" title=""Colours"" aria-label=""Colours"">';
 if(D.theme<0)h+='<option value=""-1""'+(st.theme<0?' selected':'')+'>'+esc(D.own.n)+'</option>';
 D.themes.forEach(function(t,i){h+='<option value=""'+i+'""'+(i===st.theme?' selected':'')+'>'+esc(t.n)+(t.d?' (dark)':'')+'</option>';});
 h+='</select><select id=""pickCards"" title=""Look of the cards"" aria-label=""Look of the cards"">';
 ['Filled colour','Top line','Side bar','Outline','Soft tint','Gradient'].forEach(function(n,i){h+='<option value=""'+i+'""'+(i===st.cards?' selected':'')+'>Cards: '+n+'</option>';});
 h+='</select><button type=""button"" id=""pickRound"">'+(st.round?'Square corners':'Rounded corners')+'</button><button type=""button"" id=""doPrint"">Print / PDF</button>';
 $('tools').innerHTML=h;
}
// the choices are kept in the address, so a link or a bookmark opens the page the same way
function remember(){try{history.replaceState(null,'','#t='+st.theme+'&c='+st.cards+'&r='+(st.round?1:0));}catch(e){}}
try{location.hash.replace(/([tcr])=(-?\d+)/g,function(_,k,v){v=+v;if(k==='t'&&v>=-1&&v<D.themes.length&&(v>=0||D.theme<0))st.theme=v;if(k==='c'&&v>=0&&v<6)st.cards=v;if(k==='r')st.round=v===1;});}catch(e){}

// ---- the numbers of one chart as a CSV file ----
function csvOf(i){
 var c=chartNow(i),a=results[i],q=function(s){s=String(s);return /["",\n\r]/.test(s)||/^[=+\-@]/.test(s)?'""'+(/^[=+\-@]/.test(s)?""'"":'')+s.replace(/""/g,'""""')+'""':s;};
 var lines=[[c.dimName].concat(a.S.map(function(s){return s.name;})).map(q).join(',')];
 a.cats.forEach(function(n,k){lines.push([q(shortLabel(c.dim,n))].concat(a.S.map(function(s){return String(s.V[k]);})).join(','));});
 return lines.join('\r\n');
}
function save(name,text){
 var blob=new Blob(['\uFEFF'+text],{type:'text/csv;charset=utf-8'}),u=URL.createObjectURL(blob),l=document.createElement('a');
 l.href=u;l.download=name;document.body.appendChild(l);l.click();l.remove();setTimeout(function(){URL.revokeObjectURL(u);},2000);
}

// ---- what the viewer does ----
document.addEventListener('click',function(e){
 var t=e.target,x;
 if(t.id==='pickRound'){st.round=!st.round;remember();applyLook();drawTools();drawAll();return;}
 if(t.id==='doPrint'){window.print();return;}
 if(t.hasAttribute&&t.hasAttribute('data-csv')){x=+t.getAttribute('data-csv');save(chartTitle(chartNow(x),results[x]).replace(/[\\\/:*?""<>|]+/g,' ')+'.csv',csvOf(x));return;}
 x=t.closest?t.closest('[data-clear]'):null;
 if(x){var f=x.getAttribute('data-clear');if(f==='*')SEL.clear();else SEL.delete(f);drawAll();return;}
 x=t.closest?t.closest('[data-v]'):null;
 if(x&&x.hasAttribute('data-f')){toggle(x.getAttribute('data-f'),x.getAttribute('data-v'));hideTip();drawAll();}
});
document.addEventListener('change',function(e){
 var t=e.target;
 if(t.id==='pickTheme'){st.theme=+t.value;remember();applyLook();drawAll();}
 else if(t.id==='pickCards'){st.cards=+t.value;remember();drawCards();}
 else if(t.hasAttribute('data-chart')){st.kind[+t.getAttribute('data-chart')]=+t.value;drawCharts();}
});
document.addEventListener('input',function(e){
 var t=e.target;
 if(!t.hasAttribute||!t.hasAttribute('data-find'))return;
 var si=t.getAttribute('data-find'),at=t.selectionStart;
 finds[si]=t.value;drawFilters();
 var again=document.querySelector('[data-find=""'+si+'""]');
 if(again){again.focus();try{again.setSelectionRange(at,at);}catch(x){}}
});
var tip=$('tip');
function hideTip(){tip.style.display='none';}
document.addEventListener('mousemove',function(e){
 var t=e.target.closest?e.target.closest('[data-tip]'):null;
 if(!t){hideTip();return;}
 tip.textContent=t.getAttribute('data-tip');
 tip.style.display='block';
 var w=tip.offsetWidth,h=tip.offsetHeight,x=e.clientX+14,y=e.clientY+14;
 if(x+w>window.innerWidth-6)x=e.clientX-w-10;
 if(y+h>window.innerHeight-6)y=e.clientY-h-10;
 tip.style.left=Math.max(4,x)+'px';tip.style.top=Math.max(4,y)+'px';
});
document.addEventListener('mouseleave',hideTip);
// on a narrow window (a phone) the filters fold away under one line; on a wide one they are always open
function foldFilters(){$('fbox').open=window.innerWidth>980||SEL.size>0;}
var timer=null,lastW=window.innerWidth;
window.addEventListener('resize',function(){if(window.innerWidth===lastW)return;var was=lastW>980;lastW=window.innerWidth;if(was!==(lastW>980))foldFilters();clearTimeout(timer);timer=setTimeout(function(){drawCards();drawCharts();},120);});
window.addEventListener('beforeprint',function(){drawCards();drawCharts();});

// ---- start ----
$('sub').textContent=D.sub+(D.left>0?'   |   only the first '+commas(String(N),false)+' rows are in this page':'');
$('foot').textContent=D.foot+'. This page carries its own data and works without a connection. Click a bar, a slice or a filter button to filter; click it again to undo.';
if(D.logo){$('logo').src='data:image/png;base64,'+D.logo;$('logo').hidden=false;}
$('side').hidden=!D.slicers.length;
foldFilters();
applyLook();drawTools();drawAll();
// for tests: the numbers the page works out
window.DocBriskDash={group:group,reduce:reduce,months:months,change:change,insights:insights,cardText:cardText,cardNote:cardNote,scalar:scalar,compact:compact,shortNum:shortNum,grouped:grouped,fixed:fixed,
 maskOf:maskOf,toggle:toggle,SEL:SEL,st:st,chartNow:chartNow,chartTitle:chartTitle,results:function(){return results;},drawAll:drawAll,keysOf:keysOf,shortLabel:shortLabel,axis:axis,csvOf:csvOf,kindsFor:kindsFor};
})();
</script>
</body>
</html>
";
    }
}
