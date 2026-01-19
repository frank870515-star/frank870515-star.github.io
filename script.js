// 核心逻辑文件

// 全局变量
let currentUser = null;
let adminData = [];
let filteredData = [];
let charts = {};
let currentExpensesView = localStorage.getItem('expensesView') || 'monthly';
let currentSortOrder = localStorage.getItem('sortOrder') || 'desc';

// 页面加载完成后初始化
window.addEventListener('DOMContentLoaded', () => {
    // 初始化语言（优先执行，避免闪烁）
    updateLanguage();
    
    // 检查登录状态
    checkLoginStatus();
    
    // 初始化事件监听
    initEventListeners();
});

// 检查登录状态
function checkLoginStatus() {
    const savedUser = localStorage.getItem('currentUser');
    if (savedUser) {
        currentUser = JSON.parse(savedUser);
        showDashboard();
    } else {
        showLoginPage();
    }
}

// 显示登录页面
function showLoginPage() {
    document.getElementById('login-page').classList.remove('hidden');
    document.getElementById('dashboard-page').classList.add('hidden');
}

// 显示仪表盘
function showDashboard() {
    document.getElementById('login-page').classList.add('hidden');
    document.getElementById('dashboard-page').classList.remove('hidden');
    
    // 根据用户权限显示/隐藏操作按钮
    if (currentUser.username === 'mxadmin') {
        document.getElementById('action-buttons').classList.remove('hidden');
    } else {
        document.getElementById('action-buttons').classList.add('hidden');
    }
    
    // 加载数据
    loadData();
    
    // 延迟初始化图表，提高首屏加载速度
    setTimeout(() => {
        initCharts();
        updateDashboard();
    }, 100);
}

// 初始化事件监听
function initEventListeners() {
    // 登录按钮
    const loginBtn = document.getElementById('login-btn');
    if (loginBtn) {
        loginBtn.addEventListener('click', handleLogin);
    }
    
    // 回车键登录
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    if (usernameInput) {
        usernameInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') handleLogin();
        });
    }
    if (passwordInput) {
        passwordInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') handleLogin();
        });
    }
    
    // 退出登录按钮
    const logoutBtn = document.getElementById('logout-btn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', handleLogout);
        console.log('退出登录按钮事件已绑定');
    } else {
        console.log('退出登录按钮未找到');
    }
    
    // 语言切换按钮
    const langSwitch = document.getElementById('lang-switch');
    if (langSwitch) {
        langSwitch.addEventListener('click', () => {
            switchLanguage();
            updateDashboard();
        });
    }
    
    // 指南折叠按钮
    const guideToggle = document.getElementById('guide-toggle');
    if (guideToggle) {
        guideToggle.addEventListener('click', toggleGuide);
    }
    
    // 下载模板按钮
    const downloadTemplateBtn = document.getElementById('download-template-btn');
    if (downloadTemplateBtn) {
        downloadTemplateBtn.addEventListener('click', () => {
            generateExcelTemplate();
            showAlert(t('template_download'));
        });
    }
    
    // 导入Excel按钮
    const importExcelBtn = document.getElementById('import-excel-btn');
    if (importExcelBtn) {
        importExcelBtn.addEventListener('click', () => {
            document.getElementById('excel-file-input').click();
        });
    }
    
    // Excel文件选择
    const excelFileInput = document.getElementById('excel-file-input');
    if (excelFileInput) {
        excelFileInput.addEventListener('change', handleExcelImport);
    }
    
    // 筛选输入
    // 排序按钮
    const sortAscBtn = document.getElementById('sort-asc-btn');
    const sortDescBtn = document.getElementById('sort-desc-btn');
    if (sortAscBtn) {
        sortAscBtn.addEventListener('click', () => {
            currentSortOrder = 'asc';
            localStorage.setItem('sortOrder', currentSortOrder);
            updateDashboard();
        });
    }
    if (sortDescBtn) {
        sortDescBtn.addEventListener('click', () => {
            currentSortOrder = 'desc';
            localStorage.setItem('sortOrder', currentSortOrder);
            updateDashboard();
        });
    }
    
    // 费用视图切换
    document.querySelectorAll('.chart-control-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.chart-control-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            currentExpensesView = e.target.dataset.view;
            localStorage.setItem('expensesView', currentExpensesView);
            updateExpensesChart();
        });
    });
    
    // 图表下载
    document.querySelectorAll('.chart-download').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const chartId = e.target.dataset.chart;
            if (charts[chartId]) {
                charts[chartId].getDataURL({ type: 'png' });
                showAlert(t('chart_download'));
            }
        });
    });
    
    // 数据下载
    const downloadAllBtn = document.getElementById('download-all-btn');
    const downloadFilteredBtn = document.getElementById('download-filtered-btn');
    if (downloadAllBtn) {
        downloadAllBtn.addEventListener('click', () => {
            downloadData(adminData);
        });
    }
    if (downloadFilteredBtn) {
        downloadFilteredBtn.addEventListener('click', () => {
            downloadData(filteredData.length > 0 ? filteredData : adminData);
        });
    }
}

// 处理登录
function handleLogin() {
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value.trim();
    
    if (!username || !password) {
        showAlert(t('error_empty'));
        return;
    }
    
    // 校验账号密码
    if ((username === 'mxadmin' && password === '123456') || 
        (username === 'hqadmin' && password === '123456')) {
        currentUser = { username, role: username === 'mxadmin' ? 'admin' : 'user' };
        localStorage.setItem('currentUser', JSON.stringify(currentUser));
        showDashboard();
    } else {
        showAlert(t('error_invalid'));
    }
}

// 处理退出登录
function handleLogout() {
    console.log('退出登录按钮被点击');
    currentUser = null;
    localStorage.removeItem('currentUser');
    showLoginPage();
    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
    console.log('退出登录完成');
}

// 切换指南显示
function toggleGuide() {
    const guideContent = document.getElementById('guide-content');
    const guideToggle = document.getElementById('guide-toggle');
    
    if (guideContent.classList.contains('hidden')) {
        guideContent.classList.remove('hidden');
        guideToggle.innerHTML = t('guide_title').replace('▼', '▲');
    } else {
        guideContent.classList.add('hidden');
        guideToggle.innerHTML = t('guide_title').replace('▲', '▼');
    }
}

// 加载数据
function loadData() {
    adminData = getDataFromLocalStorage();
    filteredData = [...adminData];
}

// 处理Excel导入
function handleExcelImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // 获取第一个工作表
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // 转换为JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            if (jsonData.length === 0) {
                showAlert('Excel文件为空');
                return;
            }
            
            // 校验列名
            const headers = Object.keys(jsonData[0]);
            const missingColumns = validateExcelColumns(headers);
            
            if (missingColumns.length > 0) {
                showAlert(t('error_columns').replace('{columns}', missingColumns.join('\n')));
                return;
            }
            
            // 校验数据格式
            const formatErrors = validateDataFormat(jsonData);
            
            if (formatErrors.length > 0) {
                showAlert(t('error_data_format').replace('{errors}', formatErrors.join('\n')));
                return;
            }
            
            // 处理数据
            const processedData = processExcelData(jsonData);
            
            // 检查重复月度
            const duplicateMonths = checkDuplicateMonths(processedData, adminData);
            
            if (duplicateMonths.length > 0) {
                showModal(
                    t('modal_title'),
                    t('confirm_overwrite').replace('{months}', duplicateMonths.join(', ')),
                    () => {
                        // 覆盖重复数据
                        importData(processedData, true);
                    },
                    () => {
                        // 取消导入
                        showAlert('导入已取消');
                    }
                );
            } else {
                // 直接导入
                importData(processedData, false);
            }
            
        } catch (error) {
            showAlert('Excel文件解析错误：' + error.message);
        }
        
        // 重置文件输入
        e.target.value = '';
    };
    
    reader.readAsArrayBuffer(file);
}

// 导入数据
function importData(data, overwrite) {
    console.log('开始导入数据，数据条数：', data.length);
    console.log('是否覆盖：', overwrite);
    
    if (overwrite) {
        // 覆盖重复数据
        const newMonths = data.map(item => item.month);
        adminData = adminData.filter(item => !newMonths.includes(item.month));
        console.log('已覆盖重复数据，剩余数据条数：', adminData.length);
    }
    
    // 添加新数据
    adminData = [...adminData, ...data];
    console.log('添加新数据后总条数：', adminData.length);
    
    // 按月份排序
    adminData.sort((a, b) => new Date(a.month) - new Date(b.month));
    console.log('数据已按月份排序');
    
    // 保存到localStorage
    saveDataToLocalStorage(adminData);
    console.log('数据已保存到localStorage');
    
    // 更新筛选数据
    filteredData = [...adminData];
    console.log('筛选数据已更新');
    
    // 延迟更新仪表盘，确保DOM已准备好
    setTimeout(() => {
        // 重新初始化图表
        initCharts();
        updateDashboard();
        console.log('仪表盘已更新');
    }, 100);
    
    // 显示成功消息
    const minMonth = adminData[0].month;
    const maxMonth = adminData[adminData.length - 1].month;
    showAlert(t('success_import')
        .replace('{count}', data.length)
        .replace('{minMonth}', minMonth)
        .replace('{maxMonth}', maxMonth)
    );
}

// 处理筛选
// 初始化图表
function initCharts() {
    // 满意度与SLA趋势图
    charts['satisfaction-sla'] = echarts.init(document.getElementById('satisfaction-sla-chart'));
    
    // 响应时长趋势图
    charts['response-time'] = echarts.init(document.getElementById('response-time-chart'));
    
    // 费用趋势图
    charts['expenses'] = echarts.init(document.getElementById('expenses-chart'));
    
    // 入境占比图
    charts['entry-ratio'] = echarts.init(document.getElementById('entry-ratio-chart'));
    
    // 窗口大小变化时重绘图表
    window.addEventListener('resize', () => {
        Object.values(charts).forEach(chart => chart.resize());
    });
}

// 更新仪表盘
function updateDashboard() {
    console.log('开始更新仪表盘，数据条数：', filteredData.length);
    
    // 按排序方式排序
    filteredData.sort((a, b) => {
        if (currentSortOrder === 'asc') {
            return new Date(a.month) - new Date(b.month);
        } else {
            return new Date(b.month) - new Date(a.month);
        }
    });
    
    // 更新概览卡片
    updateOverviewCards();
    console.log('概览卡片已更新');
    
    // 更新图表（按升序排列，确保2026-01在最左侧）
    const sortedData = [...filteredData].sort((a, b) => new Date(a.month) - new Date(b.month));
    updateSatisfactionSLAChart(sortedData);
    console.log('满意度与SLA图表已更新');
    
    updateResponseTimeChart(sortedData);
    console.log('响应时长图表已更新');
    
    updateExpensesChart(sortedData);
    console.log('费用图表已更新');
    
    updateEntryRatioChart(sortedData);
    console.log('入境占比图表已更新');
    
    // 更新表格
    updateDataTable();
    console.log('数据表格已更新');
}

// 更新概览卡片
function updateOverviewCards() {
    if (filteredData.length === 0) {
        document.getElementById('satisfaction-score').textContent = '-';
        document.getElementById('sla-compliance').textContent = '-';
        document.getElementById('entry-rate').textContent = '-';
        document.getElementById('total-expenses').textContent = '-';
        return;
    }
    
    // 获取最新数据
    const latestData = filteredData[0];
    
    // 满意度评分
    document.getElementById('satisfaction-score').textContent = formatNumber(latestData.satisfactionScore, 1);
    
    // SLA合格率
    document.getElementById('sla-compliance').textContent = formatPercentage(latestData.slaCompliance);
    
    // 成功入境占比
    const entryRate = (latestData.successEmployees / latestData.totalEmployees) * 100;
    document.getElementById('entry-rate').textContent = formatPercentage(entryRate);
    
    // 四大费用合计
    const totalExpenses = latestData.accommodationExpenses + 
        latestData.vehicleExpenses + 
        latestData.cateringExpenses + 
        latestData.officeExpenses;
    document.getElementById('total-expenses').textContent = formatCurrency(totalExpenses);
}

// 更新满意度与SLA趋势图
function updateSatisfactionSLAChart(data) {
    const chartData = data || filteredData;
    console.log('更新满意度与SLA图表，数据条数：', chartData.length);
    
    // 检查图表是否存在
    if (!charts['satisfaction-sla']) {
        console.log('满意度与SLA图表未初始化，重新初始化');
        charts['satisfaction-sla'] = echarts.init(document.getElementById('satisfaction-sla-chart'));
    }
    
    const chart = charts['satisfaction-sla'];
    
    const months = chartData.map(item => item.month);
    const satisfactionScores = chartData.map(item => item.satisfactionScore);
    const slaCompliances = chartData.map(item => item.slaCompliance);
    
    console.log('月度数据：', months);
    console.log('满意度数据：', satisfactionScores);
    console.log('SLA数据：', slaCompliances);
    
    const option = {
        tooltip: {
            trigger: 'axis',
            axisPointer: { type: 'cross' }
        },
        legend: {
            data: [t('satisfaction_score'), t('sla_compliance')]
        },
        grid: {
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
        },
        xAxis: {
            type: 'category',
            data: months,
            axisLabel: { rotate: 45 }
        },
        yAxis: [
            {
                type: 'value',
                name: t('satisfaction_score'),
                min: 0,
                max: 10,
                interval: 2
            },
            {
                type: 'value',
                name: t('sla_compliance'),
                min: 0,
                max: 100,
                interval: 20
            }
        ],
        series: [
            {
                name: t('satisfaction_score'),
                type: 'line',
                data: satisfactionScores,
                yAxisIndex: 0,
                smooth: true,
                lineStyle: { color: '#3b82f6' },
                itemStyle: { color: '#3b82f6' }
            },
            {
                name: t('sla_compliance'),
                type: 'line',
                data: slaCompliances,
                yAxisIndex: 1,
                smooth: true,
                lineStyle: { color: '#10b981' },
                itemStyle: { color: '#10b981' }
            }
        ]
    };
    
    chart.setOption(option);
    console.log('满意度与SLA图表已更新');
}

// 更新响应时长趋势图
function updateResponseTimeChart(data) {
    const chartData = data || filteredData;
    console.log('更新响应时长图表，数据条数：', chartData.length);
    
    // 检查图表是否存在
    if (!charts['response-time']) {
        console.log('响应时长图表未初始化，重新初始化');
        charts['response-time'] = echarts.init(document.getElementById('response-time-chart'));
    }
    
    const chart = charts['response-time'];
    
    const months = chartData.map(item => item.month);
    const responseTimes = chartData.map(item => item.responseTime);
    
    console.log('月度数据：', months);
    console.log('响应时长数据：', responseTimes);
    
    const option = {
        tooltip: {
            trigger: 'axis',
            formatter: (params) => {
                const data = params[0];
                return `${data.name}<br/>${data.seriesName}: ${data.value} ${t('response_time').split('(')[1].replace(')', '')}`;
            }
        },
        legend: {
            data: [t('response_time')]
        },
        grid: {
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
        },
        xAxis: {
            type: 'category',
            data: months,
            axisLabel: { rotate: 45 }
        },
        yAxis: {
            type: 'value',
            name: t('response_time').split('(')[1].replace(')', ''),
            min: 0
        },
        series: [
            {
                name: t('response_time'),
                type: 'line',
                data: responseTimes,
                smooth: true,
                lineStyle: { color: '#ef4444' },
                itemStyle: { color: '#ef4444' },
                markLine: {
                    data: [
                        {
                            type: 'average',
                            name: t('response_time') + ' ' + t('sla_compliance').split('(')[0]
                        }
                    ]
                }
            }
        ]
    };
    
    chart.setOption(option);
    console.log('响应时长图表已更新');
}

// 更新费用趋势图
function updateExpensesChart(data) {
    const chartData = data || filteredData;
    console.log('更新费用趋势图表，数据条数：', chartData.length);
    
    // 检查图表是否存在
    if (!charts['expenses']) {
        console.log('费用趋势图表未初始化，重新初始化');
        charts['expenses'] = echarts.init(document.getElementById('expenses-chart'));
    }
    
    const chart = charts['expenses'];
    
    // 按月份升序排序（确保2026-01在最左侧）
    const sortedData = [...chartData].sort((a, b) => new Date(a.month) - new Date(b.month));
    
    const months = sortedData.map(item => item.month);
    const accommodation = sortedData.map(item => item.accommodationExpenses);
    const vehicle = sortedData.map(item => item.vehicleExpenses);
    const catering = sortedData.map(item => item.cateringExpenses);
    const office = sortedData.map(item => item.officeExpenses);
    
    console.log('月度数据（已排序）：', months);
    
    const option = {
        tooltip: {
            trigger: 'axis',
            axisPointer: { type: 'shadow' }
        },
        legend: {
            data: [t('accommodation'), t('vehicle'), t('catering'), t('office')]
        },
        grid: {
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
        },
        xAxis: {
            type: 'category',
            data: months,
            axisLabel: { 
                rotate: 45,
                interval: 0,  // 显示所有月份
                fontSize: 11
            }
        },
        yAxis: {
            type: 'value',
            name: 'MXN'
        },
        series: [
            {
                name: t('accommodation'),
                type: currentExpensesView === 'monthly' ? 'bar' : 'line',
                data: accommodation,
                itemStyle: { color: '#3b82f6' },
                smooth: currentExpensesView === 'yearly'
            },
            {
                name: t('vehicle'),
                type: currentExpensesView === 'monthly' ? 'bar' : 'line',
                data: vehicle,
                itemStyle: { color: '#10b981' },
                smooth: currentExpensesView === 'yearly'
            },
            {
                name: t('catering'),
                type: currentExpensesView === 'monthly' ? 'bar' : 'line',
                data: catering,
                itemStyle: { color: '#f59e0b' },
                smooth: currentExpensesView === 'yearly'
            },
            {
                name: t('office'),
                type: currentExpensesView === 'monthly' ? 'bar' : 'line',
                data: office,
                itemStyle: { color: '#8b5cf6' },
                smooth: currentExpensesView === 'yearly'
            }
        ]
    };
    
    chart.setOption(option);
    console.log('费用趋势图表已更新');
}

// 更新成功入境占比趋势图
function updateEntryRatioChart(data) {
    const chartData = data || filteredData;
    console.log('更新成功入境占比趋势图表，数据条数：', chartData.length);
    
    // 检查图表是否存在
    if (!charts['entry-ratio']) {
        console.log('成功入境占比趋势图表未初始化，重新初始化');
        charts['entry-ratio'] = echarts.init(document.getElementById('entry-ratio-chart'));
    }
    
    const chart = charts['entry-ratio'];
    
    if (chartData.length === 0) {
        chart.setOption({
            series: [{
                data: []
            }]
        });
        return;
    }
    
    // 按月份升序排序（确保2026-01在最左侧）
    const sortedData = [...chartData].sort((a, b) => new Date(a.month) - new Date(b.month));
    
    const months = sortedData.map(item => item.month);
    const successRates = sortedData.map(item => {
        if (item.totalEmployees === 0) return 0;
        return (item.successEmployees / item.totalEmployees) * 100;
    });
    
    console.log('月度数据（已排序）：', months);
    console.log('成功入境占比数据：', successRates);
    
    const option = {
        tooltip: {
            trigger: 'axis',
            formatter: (params) => {
                const data = params[0];
                const index = data.dataIndex;
                const monthData = sortedData[index];
                return `${data.name}<br/>` +
                       `总计外派人数: ${monthData.totalEmployees}人<br/>` +
                       `成功入境人数: ${monthData.successEmployees}人<br/>` +
                       `${data.seriesName}: ${data.value.toFixed(1)}%`;
            }
        },
        legend: {
            data: [t('entry_rate')]
        },
        grid: {
            left: '3%',
            right: '4%',
            bottom: '3%',
            containLabel: true
        },
        xAxis: {
            type: 'category',
            data: months,
            axisLabel: { 
                rotate: 45,
                interval: 0,
                fontSize: 11
            }
        },
        yAxis: {
            type: 'value',
            name: '%',
            min: 0,
            max: 100,
            interval: 20
        },
        series: [
            {
                name: t('entry_rate'),
                type: 'line',
                data: successRates,
                smooth: true,
                lineStyle: { color: '#10b981', width: 3 },
                itemStyle: { color: '#10b981' },
                markLine: {
                    data: [
                        { 
                            type: 'average', 
                            name: '平均占比',
                            lineStyle: { color: '#f59e0b' }
                        }
                    ]
                }
            }
        ]
    };
    
    chart.setOption(option);
    console.log('成功入境占比趋势图表已更新');
}

// 更新数据表格
function updateDataTable() {
    const tbody = document.querySelector('#data-table tbody');
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" data-lang="no_data">暂无数据</td></tr>';
        return;
    }
    
    // 按月份升序排序（2026-01在最上方）
    const sortedData = [...filteredData].sort((a, b) => new Date(a.month) - new Date(b.month));
    
    const rows = sortedData.map(item => `
        <tr>
            <td>${item.month}</td>
            <td>${formatNumber(item.satisfactionScore, 1)}</td>
            <td>${formatNumber(item.responseTime, 1)}</td>
            <td>${formatCurrency(item.accommodationExpenses)}</td>
            <td>${formatCurrency(item.vehicleExpenses)}</td>
            <td>${formatCurrency(item.cateringExpenses)}</td>
            <td>${formatCurrency(item.officeExpenses)}</td>
            <td>${formatPercentage(item.slaCompliance)}</td>
            <td>${item.totalEmployees}</td>
            <td>${item.successEmployees}</td>
        </tr>
    `).join('');
    
    tbody.innerHTML = rows;
}

// 下载数据
function downloadData(data) {
    if (data.length === 0) {
        showAlert('暂无数据可下载');
        return;
    }
    
    const format = document.getElementById('download-format').value;
    const filename = generateDownloadFileName(format);
    
    // 转换为下载格式
    const downloadData = data.map(item => ({
        '月度': item.month,
        '行政服务满意度评分': item.satisfactionScore,
        '工单首次响应时长(小时)': item.responseTime,
        '住宿费用(MXN)': item.accommodationExpenses,
        '用车费用(MXN)': item.vehicleExpenses,
        '餐饮费用(MXN)': item.cateringExpenses,
        '办公费用(MXN)': item.officeExpenses,
        '供应商SLA合格率(%)': item.slaCompliance,
        '月度入境墨西哥员工总人数': item.totalEmployees,
        '成功入境墨西哥员工人数': item.successEmployees
    }));
    
    if (format === 'xlsx') {
        downloadExcel(downloadData, filename);
    } else {
        downloadCSV(downloadData, filename);
    }
    
    showAlert(t('data_download'));
}