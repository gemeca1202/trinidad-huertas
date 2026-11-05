// ---------- Street geometry (OpenStreetMap), units 1e-5 deg offset from 17.05 N, -96.71 W ----------
const RAW = `Huerto Los Olivos|289,-604;303,-614;313,-619
Huerto Los Olmos|20,-664;-2,-544;-23,-429
Huerto Los Framboyanes|-24,-670;-47,-552;-59,-481;-67,-436;-81,-371
Huerto Limonares|139,-339;150,-403;161,-468;176,-539;163,-601;157,-644
Huerto Las Flores|-117,-680;-125,-632;-132,-587;-137,-565
Huerto Los Laureles|65,-658;42,-539;31,-486;20,-422;7,-356
Huerto Naranjos|202,-640;220,-547;207,-464;197,-396;191,-365
Privada de Limonares|31,-486;161,-468
Huerto Los Sabinos|110,-652;85,-534
Huerto Los Manzanos|-59,-481;-151,-494
Av. Trinidad de las Huertas|-137,-565;-93,-560;-47,-552;-2,-544;42,-539;85,-534;131,-530;176,-539;220,-547;269,-557
Huerto Los Mangos|-71,-675;-93,-560;-100,-518
Huerto Los Platanares|241,-389;197,-396;150,-403;109,-409;64,-416;20,-422;-23,-429;-67,-436;-116,-444;-158,-450
Huerto Los Olivos|289,-604;285,-621;276,-632
Huerto Los Olivos|269,-557;271,-563;279,-594
Huerto Las Flores|-137,-565;-142,-541;-151,-494;-158,-450;-170,-392;-171,-384
Huerto Los Olivos|269,-557;253,-460;241,-389;235,-331;234,-325
Privada|-125,-632;-177,-639
Privada|-132,-587;-183,-595
Huerto Los Olivos|279,-594;289,-604
Huerto Los Ciruelos|-255,-692;-201,-687;-150,-683;-117,-680
Privada|64,-416;59,-383
Privada de Limonares|161,-468;207,-464;253,-460
Privada|-116,-444;-121,-413
Privada|109,-409;105,-378
Privada|-23,-429;-24,-426;-30,-396
Privada|-158,-450;-209,-458
Privada|-151,-494;-202,-502
Huerto Los Ciruelos|-117,-680;-71,-675;-24,-670;20,-664;65,-658;110,-652;157,-644;202,-640
Av. Universidad|-171,-384;-81,-371;7,-356;139,-339;224,-328;235,-331`;
const K = Math.cos(17.05*Math.PI/180);
const toW = (la,lo)=>[lo*K*1.1132, -la*1.1057];              // units -> world meters
const llToW = (lat,lon)=>toW((lat-17.05)*1e5,(lon+96.71)*1e5);
const wToLL = (x,y)=>[17.05 + (-y/1.1057)*1e-5, -96.71 + (x/(K*1.1132))*1e-5];
const WAYS = RAW.split("\n").map(l=>{const [n,p]=l.split("|");return {n,pts:p.split(";").map(q=>{const [a,b]=q.split(",").map(Number);return toW(a,b);})};});
const LABELS = [
 ["Av. Trinidad de las Huertas",[-47,-552],[131,-530],.5,true],
 ["Av. Universidad",[7,-356],[139,-339],.5,true],
 ["Huerto Los Laureles",[31,-486],[20,-422],.5],["Huerto Limonares",[161,-468],[150,-403],.55],
 ["Huerto Naranjos",[207,-464],[197,-396],.6],["Huerto Los Olivos",[253,-460],[241,-389],.5],
 ["Huerto Los Sabinos",[110,-652],[85,-534],.3],["Huerto Los Framboyanes",[-47,-552],[-59,-481],.5],
 ["Huerto Los Olmos",[-2,-544],[-23,-429],.45],["Huerto Los Manzanos",[-59,-481],[-151,-494],.5],
 ["Huerto Los Mangos",[-71,-675],[-93,-560],.5],["Huerto Las Flores",[-151,-494],[-158,-450],.5],
 ["Huerto Los Platanares",[109,-409],[64,-416],.5],["Huerto Los Ciruelos",[-24,-670],[65,-658],.5],
 ["Priv. de Limonares",[161,-468],[31,-486],.4]
].map(([t,a,b,f,main])=>{const A=toW(...a),B=toW(...b);return {t,main:!!main,x:A[0]+(B[0]-A[0])*f,y:A[1]+(B[1]-A[1])*f,ang:Math.atan2(B[1]-A[1],B[0]-A[0])};});