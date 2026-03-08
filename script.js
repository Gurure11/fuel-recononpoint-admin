const API_URL = "https://script.google.com/macros/s/AKfycbzf61emivheEIeZ9eYFD0Z9IKVrn53QiSSTFnq6MpwvvEv0XP1Y5_TZQlwpZv12ZOyx/exec";

let reports = [];

fetch(API_URL)
.then(res => res.json())
.then(data => {

reports = data;

populateTable(data);
updateStats(data);
buildCharts(data);

});

function populateTable(data){

const table = document.querySelector("#reportTable tbody");

table.innerHTML="";

data.forEach(r=>{

let row = `
<tr>
<td>${r.ReportID}</td>
<td>${r.Date}</td>
<td>${r.Driver}</td>
<td>${r.Vehicle}</td>
<td>${r.AccidentType}</td>
<td>${r.Status}</td>
</tr>
`;

table.innerHTML += row;

});

}

function updateStats(data){

document.getElementById("totalReports").innerText = data.length;

const vehicles = new Set(data.map(r=>r.Vehicle));
document.getElementById("vehicles").innerText = vehicles.size;

const today = new Date().toISOString().slice(0,10);

const todayCount = data.filter(r=>r.Date==today).length;

document.getElementById("today").innerText = todayCount;

const resolved = data.filter(r=>r.Status=="Resolved").length;

document.getElementById("resolved").innerText = resolved;

}

function buildCharts(data){

let types = {};

data.forEach(r=>{
types[r.AccidentType] = (types[r.AccidentType]||0)+1;
});

new Chart(document.getElementById("accidentType"),{

type:"pie",

data:{
labels:Object.keys(types),
datasets:[{
data:Object.values(types)
}]
}

});

}

document.getElementById("search").addEventListener("keyup",function(){

const value = this.value.toLowerCase();

const filtered = reports.filter(r=>
r.Driver.toLowerCase().includes(value) ||
r.Vehicle.toLowerCase().includes(value) ||
r.AccidentType.toLowerCase().includes(value)
);

populateTable(filtered);

});
