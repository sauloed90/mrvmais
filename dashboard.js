fetch('dados.xlsx')
.then(response => response.arrayBuffer())
.then(data => {

const workbook = XLSX.read(data);

const sheet =
workbook.Sheets[workbook.SheetNames[0]];

const dados =
XLSX.utils.sheet_to_json(sheet);

carregarDashboard(dados);

});

function carregarDashboard(dados){

let total = 0;

dados.forEach(item => {
total += Number(item.Valor);
});

document.getElementById("totalVendas")
.innerHTML =
total.toLocaleString(
'pt-BR',
{
style:'currency',
currency:'BRL'
});

const regioes =
[...new Set(dados.map(x => x.Região))];

const filtro =
document.getElementById("filtroRegiao");

regioes.forEach(r => {

const opcao =
document.createElement("option");

opcao.value = r;
opcao.text = r;

filtro.appendChild(opcao);

});

preencherTabela(dados);

criarGrafico(dados);

filtro.addEventListener("change",()=>{

const valor = filtro.value;

const baseFiltrada =
valor === ""
? dados
: dados.filter(x => x.Região === valor);

preencherTabela(baseFiltrada);

criarGrafico(baseFiltrada);

});

}

function preencherTabela(dados){

const tbody =
document.querySelector("#tabela tbody");

tbody.innerHTML = "";

dados.forEach(item=>{

tbody.innerHTML += `
<tr>
<td>${item.Vendedor}</td>
<td>${item.Região}</td>
<td>${item.Produto}</td>
<td>${item.Valor}</td>
</tr>
`;

});

}

let chart;

function criarGrafico(dados){

const vendedores = {};
    
dados.forEach(item=>{

if(!vendedores[item.Vendedor]){
vendedores[item.Vendedor] = 0;
}

vendedores[item.Vendedor] +=
Number(item.Valor);

});

const labels =
Object.keys(vendedores);

const valores =
Object.values(vendedores);

if(chart){
chart.destroy();
}

chart =
new Chart(
document.getElementById("graficoVendas"),
{
type:'bar',
data:{
labels:labels,
datasets:[{
label:'Faturamento',
data:valores
}]
}
});
}