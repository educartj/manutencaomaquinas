<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Master das Máquinas | Dashboard Inteligente</title>

<!-- Tailwind -->
<script src="https://cdn.tailwindcss.com"></script>

<!-- XLSX -->
<script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>

<!-- Chart.js -->
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<!-- FontAwesome -->
<link rel="stylesheet"
href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css"/>

<style>

*{
  margin:0;
  padding:0;
  box-sizing:border-box;
}

body{
  background:
  radial-gradient(circle at top left,#172554 0%,#020617 45%,#000 100%);
  color:white;
  min-height:100vh;
  font-family:'Segoe UI',sans-serif;
  overflow-x:hidden;
}

/* SCROLL */
::-webkit-scrollbar{
  width:8px;
  height:8px;
}

::-webkit-scrollbar-thumb{
  background:#334155;
  border-radius:999px;
}

/* GLASS */
.glass{
  background:rgba(255,255,255,0.04);
  border:1px solid rgba(255,255,255,0.08);
  backdrop-filter:blur(18px);
}

/* CARD */
.card{
  transition:0.3s;
}

.card:hover{
  transform:translateY(-5px);
  box-shadow:0 0 30px rgba(56,189,248,0.18);
}

/* TITLE */
.gradient-text{
  background:linear-gradient(
    to right,
    #38bdf8,
    #22c55e,
    #a855f7
  );

  -webkit-background-clip:text;
  -webkit-text-fill-color:transparent;
}

/* KPI */
.kpi-number{
  font-size:54px;
  font-weight:900;
  margin-top:10px;
}

/* BUTTON */
.action-btn{
  transition:0.3s;
}

.action-btn:hover{
  transform:translateY(-3px) scale(1.02);
}

/* TABLE */
.table-line{
  transition:0.2s;
}

.table-line:hover{
  background:rgba(255,255,255,0.05);
}

/* BADGE */
.badge{
  padding:6px 14px;
  border-radius:999px;
  font-size:12px;
  font-weight:700;
}

/* CRONOGRAMA */
.schedule-day{
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:20px;
  padding:18px;
  border-radius:18px;
  background:rgba(255,255,255,0.03);
  border:1px solid rgba(255,255,255,0.04);
  transition:0.3s;
}

.schedule-day:hover{
  background:rgba(255,255,255,0.06);
}

.machine-qty{
  background:rgba(56,189,248,0.12);
  border:1px solid rgba(56,189,248,0.18);
  color:#38bdf8;
  padding:6px 12px;
  border-radius:999px;
  font-size:12px;
  font-weight:700;
}

.update-row{
  background:rgba(250,204,21,0.08);
  border:1px solid rgba(250,204,21,0.2);
}

.discipline-name{
  color:#cbd5e1;
  font-weight:700;
}

/* PAGINATION */
.pagination-btn{
  transition:0.3s;
}

.pagination-btn:hover{
  transform:translateY(-2px);
}

</style>
</head>

<body class="p-6">

<!-- ========================================= -->
<!-- HEADER -->
<!-- ========================================= -->

<div class="flex flex-col xl:flex-row justify-between items-start xl:items-center gap-6 mb-10">

  <div class="flex items-center gap-5">

    <div class="bg-cyan-500/20 p-5 rounded-3xl">

      <i class="fa-solid fa-screwdriver-wrench text-cyan-400 text-5xl"></i>

    </div>

    <div>

      <h1 class="text-6xl font-black gradient-text">
        MASTER DAS MÁQUINAS
      </h1>

      <p class="text-slate-400 text-lg mt-3">
        Dashboard Inteligente de Controle Operacional
      </p>

    </div>

  </div>

  <!-- BUTTONS -->
  <div class="flex flex-wrap gap-4">

    <!-- UPLOAD -->
    <label class="cursor-pointer">

      <input
        type="file"
        id="upload"
        class="hidden"
        accept=".xlsx,.xls"
      >

      <div class="action-btn bg-gradient-to-r from-cyan-500 to-blue-600 px-8 py-5 rounded-3xl font-bold shadow-2xl">

        <i class="fa-solid fa-upload mr-3"></i>

        Upload da Planilha

      </div>

    </label>

  </div>

</div>

<!-- ========================================= -->
<!-- KPI -->
<!-- ========================================= -->

<div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6 mb-10">

  <div class="glass card rounded-3xl p-7">

    <div class="flex justify-between items-center">

      <div>

        <div class="text-slate-400">
          Total Equipamentos
        </div>

        <div
          id="totalEquipamentos"
          class="kpi-number text-cyan-400"
        >
          0
        </div>

      </div>

      <div class="bg-cyan-500/20 p-5 rounded-2xl">

        <i class="fa-solid fa-toolbox text-cyan-400 text-3xl"></i>

      </div>

    </div>

  </div>

  <div class="glass card rounded-3xl p-7">

    <div class="flex justify-between items-center">

      <div>

        <div class="text-slate-400">
          Disciplinas
        </div>

        <div
          id="totalDisciplinas"
          class="kpi-number text-green-400"
        >
          0
        </div>

      </div>

      <div class="bg-green-500/20 p-5 rounded-2xl">

        <i class="fa-solid fa-layer-group text-green-400 text-3xl"></i>

      </div>

    </div>

  </div>

  <div class="glass card rounded-3xl p-7">

    <div class="flex justify-between items-center">

      <div>

        <div class="text-slate-400">
          Em Manutenção
        </div>

        <div
          id="emManutencao"
          class="kpi-number text-yellow-400"
        >
          0
        </div>

      </div>

      <div class="bg-yellow-500/20 p-5 rounded-2xl">

        <i class="fa-solid fa-screwdriver text-yellow-400 text-3xl"></i>

      </div>

    </div>

  </div>

  <div class="glass card rounded-3xl p-7">

    <div class="flex justify-between items-center">

      <div>

        <div class="text-slate-400">
          Concluídos
        </div>

        <div
          id="concluido"
          class="kpi-number text-red-400"
        >
          0
        </div>

      </div>

      <div class="bg-red-500/20 p-5 rounded-2xl">

        <i class="fa-solid fa-circle-check text-red-400 text-3xl"></i>

      </div>

    </div>

  </div>

</div>

<!-- ========================================= -->
<!-- GRÁFICOS -->
<!-- ========================================= -->

<div class="grid grid-cols-1 xl:grid-cols-2 gap-6 mb-10">

  <div class="glass rounded-3xl p-7">

    <h2 class="text-3xl font-black mb-6">
      Status dos Equipamentos
    </h2>

    <canvas id="statusChart"></canvas>

  </div>

  <div class="glass rounded-3xl p-7">

    <h2 class="text-3xl font-black mb-6">
      Quantidade por Disciplina
    </h2>

    <canvas id="disciplinaChart"></canvas>

  </div>

</div>

<!-- ========================================= -->
<!-- DISCIPLINAS -->
<!-- ========================================= -->

<div class="glass rounded-3xl p-7 mb-10">

  <div class="flex items-center gap-4 mb-8">

    <div class="bg-cyan-500/20 p-4 rounded-2xl">

      <i class="fa-solid fa-chart-pie text-cyan-400 text-2xl"></i>

    </div>

    <div>

      <h2 class="text-4xl font-black">
        Disciplinas
      </h2>

      <p class="text-slate-400 mt-2">
        Quantidade de máquinas por disciplina
      </p>

    </div>

  </div>

  <div
    id="disciplinasContainer"
    class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-5"
  ></div>

</div>

<!-- ========================================= -->
<!-- CRONOGRAMA -->
<!-- ========================================= -->

<div class="glass rounded-3xl p-7 mb-10">

  <div class="flex flex-col xl:flex-row justify-between gap-5 items-start xl:items-center mb-10">

    <div>

      <h2 class="text-4xl font-black">
        Cronograma de Busca de Máquinas
      </h2>

      <p class="text-slate-400 mt-3">
        Planejamento automático baseado na planilha importada
      </p>

    </div>

    <div class="glass rounded-2xl px-6 py-4">

      <div class="text-slate-400 text-sm">
        Operação
      </div>

      <div class="text-2xl font-black text-cyan-400">
        Junho 2026
      </div>

    </div>

  </div>

  <div
    id="cronogramaContainer"
    class="grid grid-cols-1 xl:grid-cols-2 gap-6"
  ></div>

</div>

<!-- ========================================= -->
<!-- TABELA -->
<!-- ========================================= -->

<div class="glass rounded-3xl p-7">

  <div class="flex flex-col xl:flex-row justify-between gap-5 mb-7">

    <div>

      <h2 class="text-4xl font-black">
        Dados da Planilha
      </h2>

      <p class="text-slate-400 mt-2">
        Visualização dinâmica dos dados importados
      </p>

    </div>

    <input
      type="text"
      id="search"
      placeholder="Pesquisar..."
      class="bg-slate-900 border border-slate-700 px-6 py-4 rounded-2xl outline-none xl:w-80"
    />

  </div>

  <div class="overflow-auto rounded-2xl">

    <table class="w-full text-sm">

      <thead id="tableHead"></thead>

      <tbody id="tableBody"></tbody>

    </table>

  </div>

  <!-- PAGINAÇÃO -->
  <div class="flex flex-col md:flex-row justify-between items-center gap-5 mt-8">

    <div class="text-slate-400">
      Exibindo 10 registros por página
    </div>

    <div class="flex items-center gap-4">

      <button
        onclick="prevPage()"
        class="pagination-btn bg-slate-800 hover:bg-slate-700 px-5 py-3 rounded-2xl"
      >

        <i class="fa-solid fa-arrow-left"></i>

      </button>

      <div
        id="pageInfo"
        class="bg-slate-900 px-6 py-3 rounded-2xl font-bold"
      >
        Página 1
      </div>

      <button
        onclick="nextPage()"
        class="pagination-btn bg-cyan-500 hover:bg-cyan-600 px-5 py-3 rounded-2xl"
      >

        <i class="fa-solid fa-arrow-right"></i>

      </button>

    </div>

  </div>

</div>

<!-- ========================================= -->
<!-- SCRIPT -->
<!-- ========================================= -->

<script>

let originalData = [];
let filteredData = [];

let currentPage = 1;
const rowsPerPage = 10;

let statusChart;
let disciplinaChart;

document
.getElementById('upload')
.addEventListener('change',handleFile);

function handleFile(e){

  const file = e.target.files[0];

  const reader = new FileReader();

  reader.onload = function(event){

    const data =
      new Uint8Array(event.target.result);

    const workbook =
      XLSX.read(data,{type:'array'});

    const sheetName =
      workbook.SheetNames[0];

    const worksheet =
      workbook.Sheets[sheetName];

    const json =
      XLSX.utils.sheet_to_json(worksheet);

    originalData = json;
    filteredData = json;

    processData(json);

    renderTable(json);

  };

  reader.readAsArrayBuffer(file);

}

function processData(data){

  if(data.length === 0) return;

  const keys = Object.keys(data[0]);

  // REFERÊNCIAS DINÂMICAS
  const statusKey =
    keys.find(k=>k.toLowerCase().includes('status'));

  const disciplinaKey =
    keys.find(k=>k.toLowerCase().includes('disciplina'));

  const statusCount = {};
  const disciplinaCount = {};

  data.forEach(row=>{

    const status =
      row[statusKey] || 'Não informado';

    statusCount[status] =
      (statusCount[status] || 0) + 1;

    const disciplina =
      row[disciplinaKey] || 'Não informado';

    disciplinaCount[disciplina] =
      (disciplinaCount[disciplina] || 0) + 1;

  });

  // KPI
  document.getElementById('totalEquipamentos')
  .innerText = data.length;

  document.getElementById('totalDisciplinas')
  .innerText = Object.keys(disciplinaCount).length;

  const manutencao =
    Object.keys(statusCount)
    .find(k=>k.toLowerCase().includes('manut'));

  const concluido =
    Object.keys(statusCount)
    .find(k=>k.toLowerCase().includes('concl'));

  document.getElementById('emManutencao')
  .innerText =
    manutencao
    ? statusCount[manutencao]
    : 0;

  document.getElementById('concluido')
  .innerText =
    concluido
    ? statusCount[concluido]
    : 0;

  renderCharts(statusCount,disciplinaCount);

  renderDisciplinas(disciplinaCount);

  generateCronograma(disciplinaCount);

}

function renderCharts(statusData,disciplinaData){

  if(statusChart) statusChart.destroy();
  if(disciplinaChart) disciplinaChart.destroy();

  statusChart = new Chart(
    document.getElementById('statusChart'),
    {
      type:'doughnut',

      data:{
        labels:Object.keys(statusData),

        datasets:[{
          data:Object.values(statusData)
        }]
      },

      options:{
        plugins:{
          legend:{
            labels:{
              color:'white'
            }
          }
        }
      }
    }
  );

  disciplinaChart = new Chart(
    document.getElementById('disciplinaChart'),
    {
      type:'bar',

      data:{
        labels:Object.keys(disciplinaData),

        datasets:[{
          label:'Quantidade',
          data:Object.values(disciplinaData)
        }]
      },

      options:{
        scales:{
          x:{
            ticks:{color:'white'}
          },

          y:{
            ticks:{color:'white'}
          }
        },

        plugins:{
          legend:{
            labels:{
              color:'white'
            }
          }
        }
      }
    }
  );

}

function renderDisciplinas(data){

  const container =
    document.getElementById('disciplinasContainer');

  container.innerHTML = '';

  Object.entries(data)
  .sort((a,b)=>b[1]-a[1])
  .forEach(([nome,quantidade])=>{

    container.innerHTML += `

      <div class="glass card rounded-3xl p-6">

        <div class="text-slate-400">
          Disciplina
        </div>

        <div class="text-2xl font-black mt-3">
          ${nome}
        </div>

        <div class="text-5xl font-black text-cyan-400 mt-5">
          ${quantidade}
        </div>

      </div>

    `;

  });

}

function generateCronograma(disciplinas){

  const container =
    document.getElementById('cronogramaContainer');

  container.innerHTML = '';

  const locais = [
    'Residencial',
    'Escritório',
    'Área Pública'
  ];

  const diasSemana = [
    'Segunda',
    'Terça',
    'Quarta',
    'Quinta'
  ];

  const cores = [
    'cyan',
    'green',
    'yellow',
    'red',
    'purple'
  ];

  // =========================================
  // TRANSFORMA DISCIPLINAS
  // =========================================

  let disciplinasArray =
    Object.entries(disciplinas)
    .map(([nome,quantidade])=>({
      nome,
      quantidade
    }));

  // =========================================
  // DISTRIBUIÇÃO INTELIGENTE
  // MAXIMO 100 POR DIA
  // MULTIPLAS DISCIPLINAS
  // =========================================

  let agenda = [];

  while(disciplinasArray.some(d=>d.quantidade > 0)){

    let diaAtual = [];
    let totalDia = 0;

    for(let disciplina of disciplinasArray){

      if(disciplina.quantidade <= 0)
        continue;

      let disponivel =
        100 - totalDia;

      if(disponivel <= 0)
        break;

      let qtd =
        disciplina.quantidade <= disponivel
        ? disciplina.quantidade
        : disponivel;

      diaAtual.push({
        nome:disciplina.nome,
        quantidade:qtd
      });

      disciplina.quantidade -= qtd;

      totalDia += qtd;
    }

    agenda.push(diaAtual);

  }

  // =========================================
  // QUANTIDADE DE SEMANAS
  // =========================================

  const semanasNecessarias =
    Math.ceil(agenda.length / 4);

  let contadorAgenda = 0;
  let dataAtual = 1;

  // =========================================
  // GERAÇÃO
  // =========================================

  for(let semana=0; semana<semanasNecessarias; semana++){

    const cor =
      cores[semana % cores.length];

    const inicioSemana =
      String(dataAtual).padStart(2,'0');

    const fimSemana =
      String(dataAtual + 4).padStart(2,'0');

    let html = `

      <div class="glass rounded-3xl p-6 border border-${cor}-500/20">

        <div class="flex justify-between items-center mb-6">

          <div>

            <div class="text-${cor}-400 text-3xl font-black">
              Semana ${semana + 1}
            </div>

            <div class="text-slate-400">
              Busca Operacional
            </div>

          </div>

          <div class="badge bg-${cor}-500">
            ${inicioSemana} - ${fimSemana}
          </div>

        </div>

        <div class="space-y-4">

    `;

    // =========================================
    // 4 DIAS OPERACIONAIS
    // =========================================

    for(let dia=0; dia<4; dia++){

      if(contadorAgenda >= agenda.length)
        break;

      const disciplinasDia =
        agenda[contadorAgenda];

      const local =
        locais[dia % locais.length];

      let disciplinasHTML = '';

      let totalDia = 0;

      disciplinasDia.forEach(item=>{

        totalDia += item.quantidade;

        disciplinasHTML += `

          <div class="flex justify-between items-center mt-2">

            <div class="discipline-name">
              ${item.nome}
            </div>

            <div class="machine-qty">
              ${item.quantidade}
            </div>

          </div>

        `;

      });

      html += `

        <div class="schedule-day">

          <div class="w-full">

            <div class="flex justify-between items-center mb-3">

              <div>

                <div class="font-bold">
                  ${diasSemana[dia]}
                  (${String(dataAtual).padStart(2,'0')})
                </div>

                <div class="text-slate-400 text-sm mt-1">
                  ${local}
                </div>

              </div>

              <div class="bg-cyan-500/15 border border-cyan-500/20 px-4 py-2 rounded-xl">

                <span class="text-cyan-400 font-bold">
                  ${totalDia} máquinas
                </span>

              </div>

            </div>

            ${disciplinasHTML}

          </div>

        </div>

      `;

      contadorAgenda++;
      dataAtual++;

    }

    // =========================================
    // SEXTA
    // =========================================

    html += `

      <div class="schedule-day update-row">

        <div>

          <div class="font-bold text-yellow-400">
            Sexta
            (${String(dataAtual).padStart(2,'0')})
          </div>

          <div class="text-slate-400 text-sm mt-1">
            Escritório
          </div>

        </div>

        <div class="text-yellow-400 font-bold">
          Atualização da Planilha
        </div>

      </div>

    `;

    html += `
        </div>
      </div>
    `;

    container.innerHTML += html;

    dataAtual += 2;

  }

}

function renderTable(data){

  const tableHead =
    document.getElementById('tableHead');

  const tableBody =
    document.getElementById('tableBody');

  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  if(data.length === 0) return;

  const cols = Object.keys(data[0]);

  let head = '<tr>';

  cols.forEach(col=>{

    head += `
      <th class="text-left p-5 border-b border-slate-700 text-slate-300 whitespace-nowrap">
        ${col}
      </th>
    `;

  });

  head += '</tr>';

  tableHead.innerHTML = head;

  updateTablePage();

}

function updateTablePage(){

  const tableBody =
    document.getElementById('tableBody');

  tableBody.innerHTML = '';

  const start =
    (currentPage - 1) * rowsPerPage;

  const end =
    start + rowsPerPage;

  const pageData =
    filteredData.slice(start,end);

  pageData.forEach(row=>{

    let tr = `
      <tr class="table-line">
    `;

    Object.values(row).forEach(value=>{

      tr += `
        <td class="p-5 border-b border-slate-800 whitespace-nowrap">
          ${value || ''}
        </td>
      `;

    });

    tr += '</tr>';

    tableBody.innerHTML += tr;

  });

  const totalPages =
    Math.ceil(filteredData.length / rowsPerPage);

  document.getElementById('pageInfo')
  .innerText =
    `Página ${currentPage} de ${totalPages}`;

}

function nextPage(){

  const totalPages =
    Math.ceil(filteredData.length / rowsPerPage);

  if(currentPage < totalPages){

    currentPage++;

    updateTablePage();

  }

}

function prevPage(){

  if(currentPage > 1){

    currentPage--;

    updateTablePage();

  }

}

document
.getElementById('search')
.addEventListener('input',function(){

  const term =
    this.value.toLowerCase();

  filteredData =
    originalData.filter(row=>

      JSON.stringify(row)
      .toLowerCase()
      .includes(term)

    );

  currentPage = 1;

  renderTable(filteredData);

});

function exportExcel(){

  if(filteredData.length === 0){

    alert('Nenhum dado carregado.');

    return;

  }

  const worksheet =
    XLSX.utils.json_to_sheet(filteredData);

  const workbook =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    'Dashboard'
  );

  XLSX.writeFile(
    workbook,
    'dashboard_maquinas.xlsx'
  );

}

</script>

</body>

