
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AI 위험성평가표 자동 생성 에이전트</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS (Excel 다운로드) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- html2pdf.js (PDF 다운로드 전용 라이브러리) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
        body {
            font-family: 'Noto Sans KR', sans-serif;
            background-color: #f1f5f9;
        }
        .table-header {
            background-color: #2b4c7e;
            color: white;
        }
        .table-subheader {
            background-color: #3b5998;
            color: white;
        }
        .risk-badge-high {
            background-color: #fee2e2;
            color: #dc2626;
            font-weight: bold;
        }
        .risk-badge-mid {
            background-color: #fef3c7;
            color: #d97706;
            font-weight: bold;
        }
        .risk-badge-low {
            background-color: #d1fae5;
            color: #059669;
            font-weight: bold;
        }
        /* PDF 출력용 전용 스타일 */
        .pdf-export-mode {
            background: white !important;
            padding: 10px !important;
            width: 1200px !important;
        }
        .pdf-export-mode input, .pdf-export-mode textarea {
            border: none !important;
            background: transparent !important;
            outline: none !important;
            resize: none !important;
        }
    </style>
</head>
<body class="p-4 md:p-6 text-slate-800">

    <div class="max-w-[1600px] mx-auto bg-white rounded-xl shadow-lg border border-slate-200 overflow-hidden">
        
        <!-- 상단 헤더 및 제어판 -->
        <div class="bg-slate-900 text-white p-5 no-print">
            <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div>
                    <h1 class="text-2xl font-bold flex items-center gap-2">
                        <i class="fa-solid fa-shield-halved text-blue-400"></i>
                        AI 위험성평가표 자동 생성 에이전트
                    </h1>
                    <p class="text-slate-400 text-sm mt-1">현장 사진을 업로드하면 AI가 위험요인을 분석하여 표준 위험성평가표를 자동 작성합니다.</p>
                </div>
                
                <!-- 내보내기 버튼 모음 -->
                <div class="flex flex-wrap gap-2">
                    <button onclick="addEmptyRow()" class="bg-slate-700 hover:bg-slate-600 text-white px-3 py-2 rounded-lg text-sm font-medium transition flex items-center gap-1.5">
                        <i class="fa-solid fa-plus"></i> 행 추가
                    </button>
                    <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded-lg text-sm font-semibold transition flex items-center gap-1.5 shadow-sm active:scale-95">
                        <i class="fa-solid fa-file-excel text-base"></i> Excel 저장
                    </button>
                    <button onclick="exportToPDF()" class="bg-rose-600 hover:bg-rose-500 text-white px-4 py-2 rounded-lg text-sm font-semibold transition flex items-center gap-1.5 shadow-sm active:scale-95">
                        <i class="fa-solid fa-file-pdf text-base"></i> PDF 저장
                    </button>
                </div>
            </div>

            <!-- 사진 업로드 및 분석 영역 -->
            <div class="mt-6 p-4 bg-slate-800/80 rounded-xl border border-slate-700">
                <div class="grid grid-cols-1 lg:grid-cols-12 gap-4 items-center">
                    <div class="lg:col-span-5">
                        <label class="block text-xs font-semibold text-slate-300 mb-2">현장 사진 업로드 (AI 분석)</label>
                        <div class="flex items-center gap-3">
                            <input type="file" id="imageInput" accept="image/*" onchange="handleImageUpload(event)" class="block w-full text-sm text-slate-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-600 file:text-white hover:file:bg-blue-500 cursor-pointer">
                        </div>
                    </div>
                    
                    <div class="lg:col-span-4">
                        <span class="block text-xs font-semibold text-slate-300 mb-2">샘플 데이터 선택</span>
                        <div class="flex gap-2">
                            <button onclick="loadSampleData('office')" class="bg-slate-700 hover:bg-slate-600 text-xs px-3 py-2 rounded border border-slate-600 transition text-slate-200">
                                <i class="fa-solid fa-building text-blue-300 mr-1"></i> 사무실 샘플
                            </button>
                            <button onclick="loadSampleData('factory')" class="bg-slate-700 hover:bg-slate-600 text-xs px-3 py-2 rounded border border-slate-600 transition text-slate-200">
                                <i class="fa-solid fa-factory text-amber-300 mr-1"></i> 현장/공장 샘플
                            </button>
                        </div>
                    </div>

                    <div class="lg:col-span-3 text-right">
                        <button id="analyzeBtn" onclick="runAIAnalysis()" disabled class="w-full bg-blue-600 hover:bg-blue-500 disabled:bg-slate-700 text-white font-bold py-2.5 px-4 rounded-lg transition flex items-center justify-center gap-2">
                            <i class="fa-solid fa-wand-magic-sparkles"></i>
                            <span id="btnText">AI 위험요인 자동 분석</span>
                        </button>
                    </div>
                </div>

                <!-- 업로드 이미지 미리보기 -->
                <div id="imagePreviewContainer" class="hidden mt-4 pt-4 border-t border-slate-700 flex items-center gap-4">
                    <img id="imagePreview" class="h-20 w-32 object-cover rounded-lg border border-slate-600" src="" alt="업로드 이미지">
                    <div class="text-xs text-slate-300">
                        <p class="font-bold text-sm text-white" id="fileNameText">이미지 로드됨</p>
                        <p class="mt-0.5">버튼을 누르면 AI가 유해위험요인, 재해유형, 대책 및 위험성 점수를 추출합니다.</p>
                    </div>
                </div>
            </div>
        </div>

        <!-- 평가표 상단 문서 정보 영역 -->
        <div id="evaluationTableArea" class="p-6 overflow-x-auto bg-white">
            <div class="mb-4 flex flex-wrap justify-between items-end border-b-2 border-slate-800 pb-3">
                <div>
                    <h2 class="text-3xl font-extrabold tracking-tight text-slate-900 text-center md:text-left">위 험 성 평 가 표</h2>
                </div>
                <div class="flex flex-wrap gap-6 text-sm font-semibold text-slate-700 mt-2 md:mt-0">
                    <div class="flex items-center gap-2">
                        <span class="text-slate-500">평가일자:</span>
                        <input type="text" id="evalDate" value="2026년 정기평가" class="border-b border-slate-400 focus:outline-none focus:border-blue-600 px-1 py-0.5 font-bold text-slate-800">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-slate-500">평가부서:</span>
                        <input type="text" id="evalDept" value="00팀" class="border-b border-slate-400 focus:outline-none focus:border-blue-600 px-1 py-0.5 font-bold text-slate-800 w-24">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="text-slate-500">평가자:</span>
                        <input type="text" id="evalUser" value="홍길동 (인)" class="border-b border-slate-400 focus:outline-none focus:border-blue-600 px-1 py-0.5 font-bold text-slate-800 w-28">
                    </div>
                </div>
            </div>

            <!-- 위험성평가 메인 테이블 -->
            <table id="riskAssessmentTable" class="w-full border-collapse border border-slate-400 text-xs min-w-[1200px]">
                <thead>
                    <tr class="table-header text-center">
                        <th rowspan="2" class="border border-slate-400 p-2 w-10">순번</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-14">소속</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-10">순</th>
                        <th colspan="2" class="border border-slate-400 p-2">유해위험요인</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-20">재해유형</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-36">현재의 안전대책</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-24">위험성<br><span class="text-[10px] font-normal">(빈도4×강도5)</span></th>
                        <th colspan="3" class="border border-slate-400 p-2">개선대책</th>
                        <th colspan="2" class="border border-slate-400 p-2 w-24">개선대책 반영 후<br>위험성(추정)</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-16">비고</th>
                        <th rowspan="2" class="border border-slate-400 p-2 w-10 no-print-col">삭제</th>
                    </tr>
                    <tr class="table-subheader text-center">
                        <th class="border border-slate-400 p-2 w-32">장소/설비/공정</th>
                        <th class="border border-slate-400 p-2">세부 위험요인</th>
                        <th class="border border-slate-400 p-2">추가 가능대책</th>
                        <th class="border border-slate-400 p-2 w-16">개선담당</th>
                        <th class="border border-slate-400 p-2 w-20">완료일</th>
                        <th class="border border-slate-400 p-2 w-10">빈도</th>
                        <th class="border border-slate-400 p-2 w-10">강도</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    <!-- Dynamic Rows -->
                </tbody>
            </table>

            <div class="mt-4 flex justify-between items-center text-xs text-slate-500">
                <p>* 빈도(1~4단계) × 강도(1~5단계) = 위험성 점수. (12점 이상: 위험성 높음, 개선대책 필수)</p>
                <p>문서 생성일: <span id="createdDate">2026-08-12</span></p>
            </div>
        </div>

    </div>

    <!-- 스크립트 로직 -->
    <script>
        const apiKey = "";
        let uploadedImageBase64 = null;

        // 샘플 데이터 (첨부한 위험성평가표 이미지 100% 매핑)
        const sampleOfficeData = [
            { id: 1, dept: "공통", seq: 1, item: "프레젠테이션 스크린/프로젝터", detail: "천장 고정 프로젝터 또는 스크린 낙하 위험", type: "낙하·비래", current: "정기 점검 실시", freq: 3, sev: 4, action: "설치 볼트·행거 연 1회 이상 점검 및 교체", owner: "시설팀", date: "2025.06", postFreq: 2, postSev: 3, remark: "" },
            { id: 2, dept: "공통", seq: 2, item: "전기설비 (콘센트·멀티탭)", detail: "콘센트 과부하 및 문어발식 배선으로 인한 화재 위험", type: "화재", current: "개인별 전기 방지기 설치", freq: 3, sev: 4, action: "멀티탭 정격용량 표시 스티커 부착 및 교육", owner: "시설팀", date: "2025.05", postFreq: 2, postSev: 3, remark: "" },
            { id: 3, dept: "공통", seq: 3, item: "의자·테이블 이동 및 재배치", detail: "가구 이동 시 허리·발 충돌 부상 위험", type: "부딪힘", current: "이동 시 보조 인력 배치", freq: 3, sev: 4, action: "이동용 손잡이 달린 테이블 도입, 이동 매뉴얼 부착", owner: "시설팀", date: "2025.07", postFreq: 2, postSev: 2, remark: "" },
            { id: 4, dept: "공통", seq: 4, item: "비상구·통로", detail: "회의 중 비상구 앞 가구 배치로 대피 지연 위험", type: "화재/기타", current: "비상구 표지판 설치", freq: 3, sev: 4, action: "비상구 전면 2m 이내 가구 배치 금지 표시 바닥 부착", owner: "시설팀", date: "2025.05", postFreq: 1, postSev: 2, remark: "" },
            { id: 5, dept: "공통", seq: 5, item: "소화기 접근성", detail: "소화기 앞 물품 적치로 초기 대응 지연 위험", type: "화재", current: "소화기 위치 표지판 설치", freq: 3, sev: 4, action: "소화기 주변 50cm 이내 물품 적치 금지 바닥 표시", owner: "시설팀", date: "2025.05", postFreq: 1, postSev: 2, remark: "" },
            { id: 6, dept: "공통", seq: 6, item: "조명 및 눈부심", detail: "프레젠테이션 중 조명 조도 부적절로 인한 눈 피로", type: "건강장해", current: "조도 기준 유지", freq: 3, sev: 4, action: "조광(디머) 스위치 또는 블라인드 설치로 조도 조절", owner: "시설팀", date: "2025.08", postFreq: 2, postSev: 2, remark: "" },
            { id: 7, dept: "공통", seq: 7, item: "단상(포디엄) 높이차", detail: "단상 오르내릴 때 발목 부상·낙상 위험", type: "넘어짐", current: "단상 경계 표시", freq: 3, sev: 4, action: "단상 앞 논슬립 테이프 부착 및 손잡이 설치 검토", owner: "시설팀", date: "2025.06", postFreq: 2, postSev: 2, remark: "" },
            { id: 8, dept: "공통", seq: 8, item: "음향·영상 장비 케이블", detail: "바닥 노출 케이블로 인한 걸려 넘어짐 위험", type: "넘어짐", current: "케이블 타이로 정리", freq: 3, sev: 4, action: "바닥 케이블 덕트 또는 케이블 커버 설치", owner: "시설팀", date: "2025.06", postFreq: 1, postSev: 2, remark: "" },
            { id: 9, dept: "공통", seq: 9, item: "냉난방 설비", detail: "에어컨 필터 미관리로 인한 실내공기질 악화", type: "건강장해", current: "정기 청소 실시", freq: 3, sev: 4, action: "분기 1회 필터 청소 점검일지 작성·관리", owner: "시설팀", date: "2025.06", postFreq: 2, postSev: 2, remark: "" }
        ];

        let currentData = [...sampleOfficeData];

        window.onload = function() {
            renderTable();
            document.getElementById('createdDate').innerText = new Date().toISOString().split('T')[0];
        };

        function renderTable() {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            currentData.forEach((row, index) => {
                const riskScore = row.freq * row.sev;
                let riskClass = 'risk-badge-low';
                if (riskScore >= 12) riskClass = 'risk-badge-high';
                else if (riskScore >= 6) riskClass = 'risk-badge-mid';

                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 transition border-b border-slate-300";
                tr.innerHTML = `
                    <td class="p-2 text-center border-r border-slate-300 font-medium">${index + 1}</td>
                    <td class="p-1 border-r border-slate-300"><input type="text" value="${row.dept}" onchange="updateCell(${index}, 'dept', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><input type="text" value="${row.seq}" onchange="updateCell(${index}, 'seq', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><input type="text" value="${row.item}" onchange="updateCell(${index}, 'item', this.value)" class="w-full font-semibold bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><textarea onchange="updateCell(${index}, 'detail', this.value)" class="w-full bg-transparent focus:bg-white border-0 resize-y text-xs p-1" rows="2">${row.detail}</textarea></td>
                    <td class="p-1 border-r border-slate-300 text-center"><input type="text" value="${row.type}" onchange="updateCell(${index}, 'type', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><textarea onchange="updateCell(${index}, 'current', this.value)" class="w-full bg-transparent focus:bg-white border-0 resize-y text-xs p-1" rows="2">${row.current}</textarea></td>
                    <td class="p-1 border-r border-slate-300 text-center ${riskClass}">
                        <div class="flex items-center justify-center gap-0.5">
                            <input type="number" min="1" max="4" value="${row.freq}" onchange="updateRisk(${index}, 'freq', this.value)" class="w-6 text-center bg-transparent font-bold">×
                            <input type="number" min="1" max="5" value="${row.sev}" onchange="updateRisk(${index}, 'sev', this.value)" class="w-6 text-center bg-transparent font-bold">=
                            <span class="font-extrabold text-sm ml-1">${riskScore}</span>
                        </div>
                    </td>
                    <td class="p-1 border-r border-slate-300"><textarea onchange="updateCell(${index}, 'action', this.value)" class="w-full bg-transparent focus:bg-white border-0 resize-y text-xs p-1 text-blue-900 font-medium" rows="2">${row.action}</textarea></td>
                    <td class="p-1 border-r border-slate-300"><input type="text" value="${row.owner}" onchange="updateCell(${index}, 'owner', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><input type="text" value="${row.date}" onchange="updateCell(${index}, 'date', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><input type="number" min="1" max="4" value="${row.postFreq}" onchange="updateCell(${index}, 'postFreq', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><input type="number" min="1" max="5" value="${row.postSev}" onchange="updateCell(${index}, 'postSev', this.value)" class="w-full text-center bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 border-r border-slate-300"><input type="text" value="${row.remark || ''}" onchange="updateCell(${index}, 'remark', this.value)" class="w-full bg-transparent focus:bg-white border-0"></td>
                    <td class="p-1 text-center no-print-col">
                        <button onclick="deleteRow(${index})" class="text-rose-500 hover:text-rose-700 px-2 py-1"><i class="fa-solid fa-trash"></i></button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function updateCell(index, field, value) {
            currentData[index][field] = value;
        }

        function updateRisk(index, field, value) {
            currentData[index][field] = parseInt(value) || 1;
            renderTable();
        }

        function addEmptyRow() {
            const newSeq = currentData.length + 1;
            currentData.push({
                id: Date.now(), dept: "공통", seq: newSeq, item: "신규 항목", detail: "위험요인 세부 내용 입력", type: "기타", current: "현재 안전대책", freq: 2, sev: 2, action: "개선대책 작성", owner: "담당부서", date: "2026.09", postFreq: 1, postSev: 1, remark: ""
            });
            renderTable();
        }

        function deleteRow(index) {
            currentData.splice(index, 1);
            renderTable();
        }

        function handleImageUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            document.getElementById('fileNameText').innerText = file.name;
            document.getElementById('imagePreviewContainer').classList.remove('hidden');

            const reader = new FileReader();
            reader.onload = function(e) {
                document.getElementById('imagePreview').src = e.target.result;
                uploadedImageBase64 = e.target.result.split(',')[1];
                document.getElementById('analyzeBtn').disabled = false;
            };
            reader.readAsDataURL(file);
        }

        function loadSampleData(type) {
            if (type === 'office') {
                currentData = [...sampleOfficeData];
            } else {
                currentData = [
                    { id: 101, dept: "생산", seq: 1, item: "프레스 기계 작업", detail: "방호장치 미작동으로 인한 손 끼임 위험", type: "감김·끼임", current: "작업자 주의 교육", freq: 4, sev: 5, action: "양수조작식 안전장치 및 인터록 센서 교체", owner: "안전팀", date: "2026.05", postFreq: 1, postSev: 3, remark: "우선 개선" },
                    { id: 102, dept: "설비", seq: 2, item: "고소 작업대(사다리)", detail: "안전대 미착용 상태 작업 중 추락 위험", type: "추락", current: "2인 1조 작업", freq: 3, sev: 5, action: "안전모·안전대 착용 의무화 및 바닥 고정대 설치", owner: "시설팀", date: "2026.06", postFreq: 1, postSev: 2, remark: "" },
                    { id: 103, dept: "물류", seq: 3, item: "지게차 운행 통로", detail: "보행자 통로 구분 미비로 인한 충돌 위험", type: "부딪힘", current: "경적 울림 후 이동", freq: 3, sev: 4, action: "보행자 전용 통로 도색 및 반사경 추가 설치", owner: "물류팀", date: "2026.05", postFreq: 1, postSev: 2, remark: "" }
                ];
            }
            renderTable();
        }

        // AI 비전 분석 기능
        async function runAIAnalysis() {
            if (!uploadedImageBase64) {
                alert("분석할 사진을 먼저 업로드해 주세요.");
                return;
            }

            const btn = document.getElementById('analyzeBtn');
            const btnText = document.getElementById('btnText');
            btn.disabled = true;
            btnText.innerText = "AI 위험요인 분석 중...";

            const promptText = `
                업로드된 작업 현장 사진을 분석하여 위험성평가 항목 1~3개를 추출해 주세요.
                반드시 아래 JSON 형식만 반환해 주세요:
                {
                    "items": [
                        {
                            "dept": "공통",
                            "seq": 1,
                            "item": "장소/설비명",
                            "detail": "세부 유해위험요인 내용",
                            "type": "재해유형",
                            "current": "현재 안전대책",
                            "freq": 3,
                            "sev": 4,
                            "action": "추가 개선대책",
                            "owner": "시설팀",
                            "date": "2026.06",
                            "postFreq": 1,
                            "postSev": 2,
                            "remark": ""
                        }
                    ]
                }
            `;

            const payload = {
                contents: [{
                    parts: [
                        { text: promptText },
                        { inlineData: { mimeType: "image/jpeg", data: uploadedImageBase64 } }
                    ]
                }],
                generationConfig: { responseMimeType: "application/json" }
            };

            let response = null;
            let delays = [1000, 2000, 4000, 8000];
            
            for (let i = 0; i <= delays.length; i++) {
                try {
                    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    });
                    if (res.ok) {
                        response = await res.json();
                        break;
                    }
                } catch (e) {}
                if (i < delays.length) await new Promise(r => setTimeout(r, delays[i]));
            }

            btn.disabled = false;
            btnText.innerText = "AI 위험요인 자동 분석";

            if (response && response.candidates && response.candidates[0]) {
                try {
                    const jsonText = response.candidates[0].content.parts[0].text;
                    const parsed = JSON.parse(jsonText);
                    if (parsed.items && parsed.items.length > 0) {
                        currentData = parsed.items;
                        renderTable();
                        alert("AI 사진 분석이 성공적으로 실행되었습니다!");
                    } else {
                        alert("이미지에서 위험 요소를 명확히 도출하지 못했습니다.");
                    }
                } catch (err) {
                    alert("데이터 처리 중 오류가 발생했습니다.");
                }
            } else {
                alert("AI 응답을 가져오지 못했습니다. 잠시 후 다시 시도해 주세요.");
            }
        }

        // EXCEL 저장 로직 (SheetJS 사용)
        function exportToExcel() {
            try {
                if (typeof XLSX === 'undefined') {
                    alert('엑셀 생성 모듈을 불러오는 중입니다. 잠시 후 다시 클릭해주세요.');
                    return;
                }

                const evalDate = document.getElementById('evalDate').value;
                const evalDept = document.getElementById('evalDept').value;
                const evalUser = document.getElementById('evalUser').value;

                const headers = [
                    ["위 험 성 평 가 표"],
                    [`평가일자: ${evalDate}`, "", "", `평가부서: ${evalDept}`, "", "", "", "", "", "", `평가자: ${evalUser}`],
                    [
                        "순번", "소속", "순", "장소/설비/공정", "세부 위험요인", "재해유형", "현재의 안전대책",
                        "위험성(빈도x강도)", "추가 개선대책", "개선담당", "완료일", "개선후 빈도", "개선후 강도", "비고"
                    ]
                ];

                const rows = currentData.map((d, i) => [
                    i + 1,
                    d.dept || "공통",
                    d.seq || (i + 1),
                    d.item || "",
                    d.detail || "",
                    d.type || "",
                    d.current || "",
                    `빈도${d.freq || 1}x강도${d.sev || 1}=${(d.freq || 1) * (d.sev || 1)}`,
                    d.action || "",
                    d.owner || "",
                    d.date || "",
                    d.postFreq || 1,
                    d.postSev || 1,
                    d.remark || ""
                ]);

                const worksheet = XLSX.utils.aoa_to_sheet([...headers, ...rows]);

                // 컬럼 폭 조절
                worksheet['!cols'] = [
                    { wch: 6 }, { wch: 8 }, { wch: 6 }, { wch: 25 }, { wch: 35 }, { wch: 12 },
                    { wch: 25 }, { wch: 18 }, { wch: 35 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }
                ];

                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "위험성평가표");
                
                const today = new Date().toISOString().split('T')[0];
                XLSX.writeFile(workbook, `위험성평가표_${today}.xlsx`);
            } catch (err) {
                alert("엑셀 저장 중 오류가 발생했습니다: " + err.message);
            }
        }

        // PDF 저장 로직 (html2pdf 사용)
        function exportToPDF() {
            try {
                if (typeof html2pdf === 'undefined') {
                    alert('PDF 생성 모듈을 불러오는 중입니다. 잠시 후 다시 클릭해주세요.');
                    return;
                }

                const element = document.getElementById('evaluationTableArea');
                
                // 삭제 컬럼 일시 숨김
                const deleteCols = document.querySelectorAll('.no-print-col');
                deleteCols.forEach(col => col.style.display = 'none');

                element.classList.add('pdf-export-mode');

                const opt = {
                    margin:       [8, 8, 8, 8],
                    filename:     `위험성평가표_${new Date().toISOString().split('T')[0]}.pdf`,
                    image:        { type: 'jpeg', quality: 0.98 },
                    html2canvas:  { scale: 2, useCORS: true, scrollX: 0, scrollY: 0 },
                    jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape' }
                };

                html2pdf().set(opt).from(element).save().then(() => {
                    element.classList.remove('pdf-export-mode');
                    deleteCols.forEach(col => col.style.display = '');
                }).catch(err => {
                    element.classList.remove('pdf-export-mode');
                    deleteCols.forEach(col => col.style.display = '');
                    alert('PDF 다운로드 중 오류가 발생했습니다.');
                });

            } catch (err) {
                alert("PDF 저장 오류: " + err.message);
            }
        }
    </script>
</body>
</html>

