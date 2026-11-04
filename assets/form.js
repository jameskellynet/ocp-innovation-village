/* Shared form engine. Reads window.FORM_CONFIG, set by each form page. */
(function(){
const C = window.FORM_CONFIG;
/* ===== Engine ===== */
const $ = id => document.getElementById(id);
$("eventLine").textContent = `${C.summit} · ${C.location} · ${C.dates}`;
$("eyebrow").textContent = `${C.summit} · ${C.program}`;
$("heading").textContent = C.heading;
$("lede").textContent = C.lede;
$("deadlineText").textContent = C.deadlineLabel;
$("tcLink").href = C.termsUrl;
$("footer").textContent = `${C.summit} · ${C.program} · Questions? Contact the OCP events team.`;
$("previewBar").hidden = !!C.endpoint;
(function(){const d=Math.ceil((new Date(C.deadline)-new Date())/864e5);
  $("countdown").textContent = d>1?`${d} days left`:d===1?"Closes today":"Closed";})();

$("categories").innerHTML = C.categories.map((c,i)=>`<label class="chip"><input type="radio" name="category" id="cat${i}" value="${c.replace(/"/g,'&quot;')}"><span>${c.replace(/&/g,'&amp;')}</span></label>`).join("");
$("timeline").innerHTML = C.timeline.map(([d,t,k])=>`<li class="${k||''}"><time>${d}</time><div>${t}</div></li>`).join("");
$("scoring").innerHTML = C.scoring.map(s=>`<div>${s.replace(/&/g,'&amp;')}</div>`).join("");
$("eligibility").innerHTML = C.eligibility.map((e,i)=>`<div class="check" data-elig><input type="checkbox" id="elig${i}"><label for="elig${i}">${e}</label></div>`).join("");

/* character counters */
document.querySelectorAll("[data-max]").forEach(f=>{
  const max=+f.dataset.max, ta=f.querySelector("textarea"), c=f.querySelector(".counter");
  const upd=()=>{const n=ta.value.length;c.textContent=`${n.toLocaleString()} / ${max.toLocaleString()}`;
    c.className="counter"+(n>max?" over":n>max*.9?" near":"");};
  ta.addEventListener("input",upd); upd();
});

/* presenters */
let pCount=0;
function addPresenter(v={}){
  const i=pCount++, el=document.createElement("div"); el.className="presenter";
  el.innerHTML=`<div class="ph"><span>Presenter</span><button type="button" class="linkbtn">Remove</button></div>
    <div class="two"><input type="text" id="pName${i}" placeholder="Full name" aria-label="Presenter name"><input type="email" id="pEmail${i}" placeholder="Email" aria-label="Presenter email"></div>
    <div class="two"><input type="text" id="pCo${i}" placeholder="Company" aria-label="Presenter company"><input type="text" id="pTitle${i}" placeholder="Job title" aria-label="Presenter job title"></div>`;
  [["pName",v.name],["pEmail",v.email],["pCo",v.company],["pTitle",v.title]].forEach(([k,val])=>{if(val)el.querySelector(`#${k}${i}`).value=val;});
  el.querySelector(".linkbtn").onclick=()=>{el.remove();renumber();};
  $("presenters").appendChild(el); renumber();
}
function renumber(){[...$("presenters").children].forEach((p,n)=>p.querySelector(".ph span").textContent=`Presenter ${n+1}`);}
$("addPresenter").onclick=()=>addPresenter();
addPresenter();

/* collect + validate */
function collect(){
  const v=id=>$(id).value.trim();
  const presenters=[...$("presenters").children].map(p=>{const [n,e,c,t]=p.querySelectorAll("input");return{name:n.value.trim(),email:e.value.trim(),company:c.value.trim(),title:t.value.trim()};}).filter(p=>p.name||p.email);
  const cat=document.querySelector('input[name=category]:checked');
  return {formId:C.formId,firstName:v("firstName"),lastName:v("lastName"),email:v("email"),company:v("company"),jobTitle:v("jobTitle"),
    title:v("title"),category:cat?cat.value:"",abstract:v("abstract"),whyDemo:v("whyDemo"),takeaway:v("takeaway"),
    equipment:v("equipment"),presenters,eligibilityConfirmed:true,consent:$("consent").checked,website:$("website").value};
}
function validate(){
  let first=null;
  document.querySelectorAll("[data-req]").forEach(f=>{
    let ok;
    if(f.dataset.radio) ok=!!f.querySelector("input:checked");
    else{const el=f.querySelector("input,textarea"),val=el.value.trim();
      ok=val.length>0&&(!f.dataset.max||val.length<=+f.dataset.max)&&(!("email" in f.dataset)||/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val));}
    f.classList.toggle("invalid",!ok); if(!ok&&!first)first=f;
  });
  document.querySelectorAll("[data-elig]").forEach(r=>{const ok=r.querySelector("input").checked;r.classList.toggle("invalid",!ok);if(!ok&&!first)first=r;});
  const cr=$("consentRow"),cok=$("consent").checked;cr.classList.toggle("invalid",!cok);if(!cok&&!first)first=cr;
  return first;
}
document.addEventListener("input",e=>{const f=e.target.closest(".invalid");if(f)f.classList.remove("invalid");});
document.addEventListener("change",e=>{const f=e.target.closest(".invalid");if(f)f.classList.remove("invalid");});

const esc=s=>String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
let data=null;
$("form").addEventListener("submit",e=>{
  e.preventDefault();
  const bad=validate();
  if(bad){$("formErr").textContent="Some answers need attention. They are marked in red.";bad.scrollIntoView({behavior:"smooth",block:"center"});return;}
  $("formErr").textContent=""; data=collect();
  const rows=[["Name",`${data.firstName} ${data.lastName}`],["Email",data.email],["Company",data.company],["Job title",data.jobTitle],
    ["Demo title",data.title],["Category",data.category],["Abstract",data.abstract],["Why a demo",data.whyDemo],["Attendee takeaway",data.takeaway],
    ["Equipment",data.equipment],["Presenters",data.presenters.length?data.presenters.map(p=>[p.name,p.title,p.company,p.email].filter(Boolean).join(", ")).join("\n"):"None listed"]];
  $("reviewList").innerHTML=rows.map(([k,v])=>`<dt>${k}</dt><dd>${esc(v)}</dd>`).join("");
  $("form").hidden=true;$("previewBar").hidden=true;$("review").hidden=false;window.scrollTo({top:$("review").offsetTop-20,behavior:"smooth"});
});
$("editBtn").onclick=()=>{$("review").hidden=true;$("form").hidden=false;$("previewBar").hidden=!!C.endpoint;};

$("confirmBtn").onclick=async()=>{
  const b=$("confirmBtn");b.disabled=true;b.textContent="Sending…";$("sendErr").textContent="";
  let ref="IV-"+Date.now().toString(36).toUpperCase().slice(-6);
  try{
    if(C.endpoint){
      const r=await fetch(C.endpoint,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(data)});
      const j=await r.json(); if(!j.ok)throw new Error(j.error||"The server did not accept the submission."); ref=j.ref||ref;
    }
    $("doneName").textContent=data.firstName;$("doneRef").textContent=ref;$("doneEmail").textContent=data.email;
    $("doneNotify").textContent=C.notifyLabel;$("demoNote").hidden=!!C.endpoint;
    $("review").hidden=true;$("done").hidden=false;window.scrollTo({top:$("done").offsetTop-20,behavior:"smooth"});
  }catch(err){$("sendErr").textContent=`Your submission was not sent: ${err.message} Please try again in a minute.`;}
  finally{b.disabled=false;b.textContent="Confirm my submission";}
};
$("againBtn").onclick=()=>{$("form").reset();$("presenters").innerHTML="";pCount=0;addPresenter();
  document.querySelectorAll("[data-max] textarea").forEach(t=>t.dispatchEvent(new Event("input")));
  $("done").hidden=true;$("form").hidden=false;$("previewBar").hidden=!!C.endpoint;window.scrollTo({top:0});};

/* sample entry for previewing (example data only) */
$("sampleBtn").onclick=()=>{
  const s={firstName:"Example",lastName:"Submitter",email:"example@company.com",company:"Example Co.",jobTitle:"Principal Engineer",
    title:"Hot-swap a liquid-cooled sled without draining the loop",
    abstract:"Attendees walk up to an OCP ORv3 rack running a live AI training load and replace a cold-plate-cooled compute sled using blind-mate quick disconnects. A live dashboard shows coolant flow, inlet/outlet temperature and pressure at the CDU while the swap happens, so visitors see that the loop stays stable and nothing leaks. We will explain the open specification behind the connectors and how multiple vendors interoperate.",
    whyDemo:"The point is how the connector feels and behaves under pressure. Seeing a hand pull a wet-loop sled with zero drips, and watching the telemetry stay flat, is far more convincing than a slide of the same numbers.",
    takeaway:"Visitors perform the swap themselves, read the live pressure and flow data, and leave knowing which OCP specification to use for interoperable quick disconnects.",
    equipment:"One half-height ORv3 rack (about 1.2 m tall, 180 kg) with an in-rack CDU, one 2×32 A 230 V feed, and a 55-inch display. Floor-standing; needs a 2 × 2 m footprint."};
  Object.entries(s).forEach(([k,v])=>{$(k).value=v;});
  $("cat1").checked=true;
  $("presenters").innerHTML="";pCount=0;addPresenter({name:"Example Presenter",email:"presenter@company.com",company:"Example Co.",title:"Thermal Architect"});
  document.querySelectorAll("[data-elig] input").forEach(c=>c.checked=true);$("consent").checked=true;
  document.querySelectorAll("[data-max] textarea").forEach(t=>t.dispatchEvent(new Event("input")));
  document.querySelectorAll(".invalid").forEach(f=>f.classList.remove("invalid"));
};

})();
