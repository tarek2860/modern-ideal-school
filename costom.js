


const customer=document.getElementById("customer");
const historyTableBody=document.querySelector("#historyTable tbody");
const showName=document.getElementById("showName");
const showMobile=document.getElementById("showMobile");
const balanceEl=document.getElementById("balance");
const totalInEl=document.getElementById("totalIn");
const totalOutEl=document.getElementById("totalOut");
const globalInEl=document.getElementById("globalIn");
const globalOutEl=document.getElementById("globalOut");
const modal=document.getElementById("modal");

let editId=null;
let db=JSON.parse(localStorage.getItem("taliKhata"))||{};
const saveDB=()=>localStorage.setItem("taliKhata",JSON.stringify(db));
const bn=n=>Number(n).toLocaleString("bn-BD");

function loadCustomers(){
 customer.innerHTML=`<option value="">-- Customer Select --</option>`;
 Object.keys(db).forEach(n=>customer.innerHTML+=`<option value="${n}">${n}</option>`);
}

function globalTotal(){
 let gin=0,gout=0;
 Object.values(db).forEach(c=>c.trans.forEach(t=>t.type==="in"?gin+=t.amount:gout+=t.amount));
 globalInEl.innerText=bn(gin);
 globalOutEl.innerText=bn(gout);
}

addCustomerBtn.onclick=()=>{
 if(!newCustomer.value)return alert("নাম দিন");
 db[newCustomer.value]={mobile:mobile.value,trans:[]};
 saveDB();loadCustomers();customer.value=newCustomer.value;
 newCustomer.value=mobile.value="";render();
};

deleteCustomerBtn.onclick=()=>{
 if(!customer.value)return;
 if(confirm("Delete করবেন?")){
  delete db[customer.value];saveDB();loadCustomers();historyTableBody.innerHTML="";globalTotal();
 }
};

inBtn.onclick=()=>saveTrans("in");
outBtn.onclick=()=>saveTrans("out");

function saveTrans(type){
 const c=customer.value;
 if(!c)return alert("Customer select করুন");
 if(!amount.value)return alert("টাকা দিন");
 if(editId){
  const t=db[c].trans.find(x=>x.id===editId);
  t.amount=+editAmount.value;t.desc=editDesc.value;t.date=editDate.value;t.type=type;
  editId=null;closeModal();
 }else{
  db[c].trans.push({id:Date.now(),type,amount:+amount.value,desc:desc.value,date:date.value||new Date().toISOString().slice(0,10)});
 }
 saveDB();amount.value=desc.value=date.value="";render();
}

customer.onchange=render;

function render(){
 const c=customer.value;if(!c)return;
 const keyword=document.getElementById("search").value.toLowerCase();

 showName.innerText="নাম: "+c;
 showMobile.innerText="📱 "+(db[c].mobile||"---");

 let bal=0,tin=0,tout=0;historyTableBody.innerHTML="";
 db[c].trans.forEach(t=>{
  const match=t.date.includes(keyword)||(t.desc||"").toLowerCase().includes(keyword)||(t.type==="in"?"received":"payment").includes(keyword);
  if(!match)return;

  if(t.type==="in"){bal+=t.amount;tin+=t.amount}else{bal-=t.amount;tout+=t.amount}
  historyTableBody.innerHTML+=`<tr>
  <td>${t.date}</td>
  <td>${t.desc||"-"}</td>
  <td>${t.type==="in"?"Received":"Payment"}</td>
  <td>৳ ${bn(t.amount)}</td>
  <td class="action-col">
    <button class="action-btn edit-btn" onclick="editTrans(${t.id})">✏</button>
    <button class="action-btn delete-btn" onclick="delTrans(${t.id})">🗑</button>
  </td></tr>`;
 });

 balanceEl.innerText=bn(bal);
 totalInEl.innerText=bn(tin);
 totalOutEl.innerText=bn(tout);
 globalTotal();
}

function editTrans(id){
 const t=db[customer.value].trans.find(x=>x.id===id);
 editId=id;editDate.value=t.date;editDesc.value=t.desc;editAmount.value=t.amount;
 modal.style.display="flex";
}

function delTrans(id){
 if(!confirm("Delete করবেন?"))return;
 db[customer.value].trans=db[customer.value].trans.filter(t=>t.id!==id);
 saveDB();render();
}

function closeModal(){modal.style.display="none";editId=null}
function saveEdit(type){saveTrans(type)}

/* PDF */
async function downloadPDF(){
 const c=customer.value;if(!c)return alert("Customer select করুন");
 const {jsPDF}=window.jspdf;const doc=new jsPDF();let y=10,bal=0;
 doc.text("Customer: "+c,10,y);y+=8;
 db[c].trans.forEach(t=>{doc.text(`${t.date}  ${t.desc||"-"}  ${t.type}  ${t.amount}`,10,y);y+=6;bal+=t.type==="in"?t.amount:-t.amount});
 y+=6;doc.text("Final Balance: "+bal+" Taka",10,y);
 doc.save(c+"_statement.pdf");
}

/* Excel */
function downloadExcel(){
 const c=customer.value;if(!c)return alert("Customer select করুন");
 const data=[["Customer",c],["Mobile",db[c].mobile||"---"],[],["Date","Description","Type","Amount"]];
 let bal=0;
 db[c].trans.forEach(t=>{data.push([t.date,t.desc||"-",t.type==="in"?"Received":"Payment",t.amount]);bal+=t.type==="in"?t.amount:-t.amount});
 data.push([]);data.push(["Final Balance","","",bal]);
 const ws=XLSX.utils.aoa_to_sheet(data);
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,ws,"Statement");
 XLSX.writeFile(wb,c+"_statement.xlsx");
}

loadCustomers();globalTotal();