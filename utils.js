// 工具函数文件

// 多语言配置
const languages = {
    zh: {
        login_title: '墨西哥工厂行政服务仪表盘',
        username: '账号',
        password: '密码',
        login: '登录',
        login_hint: '墨西哥行政人员账号: mxadmin / 123456\n总部员工账号: hqadmin / 123456',
        dashboard_title: '墨西哥工厂行政服务仪表盘',
        lang_switch: '切换语言',
        logout: '退出登录',
        guide_title: '使用指南 ▼',
        guide_mxadmin_title: '墨西哥工厂行政人员操作指南',
        guide_mxadmin_step1: '点击「Excel模板下载」按钮获取空白模板',
        guide_mxadmin_step2: '按照模板要求填写月度行政数据',
        guide_mxadmin_step3: '点击「Excel文件导入」按钮上传填写好的Excel文件',
        guide_mxadmin_step4: '系统会自动校验数据格式，如有错误请根据提示修正',
        guide_mxadmin_step5: '导入成功后，仪表盘会自动刷新显示最新数据',
        guide_hqadmin_title: '总部员工操作指南',
        guide_hqadmin_step1: '登录后即可查看实时行政数据仪表盘',
        guide_hqadmin_step2: '使用筛选功能查看特定时间段的数据',
        guide_hqadmin_step3: '点击图表右上角的「保存为图片」按钮下载图表',
        guide_hqadmin_step4: '点击「全量数据下载」或「筛选后数据下载」按钮下载数据文件',
        guide_hqadmin_step5: '支持Excel和CSV两种格式下载',
        download_template: 'Excel模板下载',
        import_excel: 'Excel文件导入',
        filter_placeholder: '请输入筛选条件（如2026）',
        filter_clear: '清除筛选',
        satisfaction_score: '满意度评分',
        sla_compliance: 'SLA合格率',
        entry_rate: '成功入境占比',
        total_expenses: '四大费用合计',
        satisfaction_sla_trend: '满意度与SLA趋势',
        download_chart: '保存为图片',
        response_time_trend: '工单响应时长趋势',
        expenses_trend: '费用趋势',
        expenses_view_monthly: '单月详情',
        expenses_view_yearly: '全年趋势',
        entry_ratio: '成功入境占比',
        data_table: '月度数据表格',
        sort_asc: '升序排序',
        sort_desc: '降序排序',
        month: '月度',
        response_time: '响应时长(小时)',
        accommodation: '住宿费用(MXN)',
        vehicle: '用车费用(MXN)',
        catering: '餐饮费用(MXN)',
        office: '办公费用(MXN)',
        total_employees: '总人数',
        success_employees: '成功人数',
        no_data: '暂无数据',
        download_all: '全量数据下载',
        download_filtered: '筛选后数据下载',
        download_format: '下载格式：',
        format_excel: 'Excel (.xlsx)',
        format_csv: 'CSV (.csv)',
        modal_title: '提示',
        confirm: '确认',
        cancel: '取消',
        error_empty: '请输入完整账号和密码',
        error_invalid: '账号或密码错误，请重新输入',
        success_login: '登录成功',
        success_import: '导入成功！共导入 {count} 条数据，月度范围：{minMonth} 至 {maxMonth}',
        error_columns: 'Excel文件列名不完整，请检查以下列是否存在：\n{columns}',
        error_data_format: '数据格式错误：\n{errors}',
        confirm_overwrite: '检测到以下月度数据已存在：{months}\n是否覆盖原有数据？',
        template_download: '模板下载成功',
        data_download: '数据下载成功',
        chart_download: '图表下载成功'
    },
    es: {
        login_title: 'Panel de Servicios Administrativos de la Fábrica de México',
        username: 'Nombre de usuario',
        password: 'Contraseña',
        login: 'Iniciar sesión',
        login_hint: 'Cuenta de administrador de México: mxadmin / 123456\nCuenta de empleado de la sede central: hqadmin / 123456',
        dashboard_title: 'Panel de Servicios Administrativos de la Fábrica de México',
        lang_switch: 'Cambiar idioma',
        logout: 'Cerrar sesión',
        guide_title: 'Guía de uso ▼',
        guide_mxadmin_title: 'Guía de operación para el personal administrativo de la fábrica de México',
        guide_mxadmin_step1: 'Haga clic en el botón "Descargar plantilla de Excel" para obtener una plantilla en blanco',
        guide_mxadmin_step2: 'Complete los datos administrativos mensuales según los requisitos de la plantilla',
        guide_mxadmin_step3: 'Haga clic en el botón "Importar archivo Excel" para cargar el archivo Excel completado',
        guide_mxadmin_step4: 'El sistema verificará automáticamente el formato de los datos. Si hay errores, corríjalos según las indicaciones',
        guide_mxadmin_step5: 'Después de una importación exitosa, el panel se actualizará automáticamente para mostrar los datos más recientes',
        guide_hqadmin_title: 'Guía de operación para los empleados de la sede central',
        guide_hqadmin_step1: 'Después de iniciar sesión, puede ver el panel de datos administrativos en tiempo real',
        guide_hqadmin_step2: 'Use la función de filtrado para ver datos de períodos de tiempo específicos',
        guide_hqadmin_step3: 'Haga clic en el botón "Guardar como imagen" en la esquina superior derecha del gráfico para descargarlo',
        guide_hqadmin_step4: 'Haga clic en el botón "Descargar todos los datos" o "Descargar datos filtrados" para descargar el archivo de datos',
        guide_hqadmin_step5: 'Se admiten dos formatos de descarga: Excel y CSV',
        download_template: 'Descargar plantilla de Excel',
        import_excel: 'Importar archivo Excel',
        filter_placeholder: 'Ingrese condiciones de filtrado (ej: 2026)',
        filter_clear: 'Limpiar filtro',
        satisfaction_score: 'Puntuación de satisfacción',
        sla_compliance: 'Cumplimiento del SLA',
        entry_rate: 'Tasa de ingreso exitoso',
        total_expenses: 'Total de los cuatro gastos principales',
        satisfaction_sla_trend: 'Tendencia de satisfacción y cumplimiento del SLA',
        download_chart: 'Guardar como imagen',
        response_time_trend: 'Tendencia del tiempo de respuesta de las solicitudes',
        expenses_trend: 'Tendencia de gastos',
        expenses_view_monthly: 'Detalles mensuales',
        expenses_view_yearly: 'Tendencia anual',
        entry_ratio: 'Proporción de ingresos exitosos',
        data_table: 'Tabla de datos mensuales',
        sort_asc: 'Ordenar ascendente',
        sort_desc: 'Ordenar descendente',
        month: 'Mes',
        response_time: 'Tiempo de respuesta (horas)',
        accommodation: 'Gastos de alojamiento (MXN)',
        vehicle: 'Gastos de vehículos (MXN)',
        catering: 'Gastos de alimentación (MXN)',
        office: 'Gastos de oficina (MXN)',
        total_employees: 'Número total de empleados',
        success_employees: 'Número de empleados exitosos',
        no_data: 'Sin datos disponibles',
        download_all: 'Descargar todos los datos',
        download_filtered: 'Descargar datos filtrados',
        download_format: 'Formato de descarga:',
        format_excel: 'Excel (.xlsx)',
        format_csv: 'CSV (.csv)',
        modal_title: 'Indicación',
        confirm: 'Confirmar',
        cancel: 'Cancelar',
        error_empty: 'Por favor, ingrese el nombre de usuario y la contraseña completos',
        error_invalid: 'Nombre de usuario o contraseña incorrectos. Por favor, inténtelo de nuevo',
        success_login: 'Inicio de sesión exitoso',
        success_import: '¡Importación exitosa! Se importaron {count} registros. Rango de meses: de {minMonth} a {maxMonth}',
        error_columns: 'Los nombres de columna del archivo Excel no están completos. Verifique si existen las siguientes columnas:\n{columns}',
        error_data_format: 'Error en el formato de los datos:\n{errors}',
        confirm_overwrite: 'Se detectaron los siguientes datos mensuales existentes: {months}\n¿Desea sobrescribir los datos existentes?',
        template_download: 'Descarga de plantilla exitosa',
        data_download: 'Descarga de datos exitosa',
        chart_download: 'Descarga de gráfico exitosa'
    }
};

// 当前语言
let currentLang = localStorage.getItem('currentLang') || 'zh';

// 获取翻译
function t(key) {
    return languages[currentLang][key] || key;
}

// 切换语言
function switchLanguage() {
    currentLang = currentLang === 'zh' ? 'es' : 'zh';
    localStorage.setItem('currentLang', currentLang);
    updateLanguage();
}

// 更新页面语言
function updateLanguage() {
    document.querySelectorAll('[data-lang]').forEach(el => {
        const key = el.getAttribute('data-lang');
        if (el.tagName === 'INPUT' && el.placeholder) {
            el.placeholder = t(key);
        } else if (el.tagName === 'BUTTON' || el.tagName === 'LABEL' || el.tagName === 'TH' || el.tagName === 'TD') {
            el.textContent = t(key);
        } else if (el.tagName === 'H1' || el.tagName === 'H2' || el.tagName === 'H3' || el.tagName === 'P') {
            el.textContent = t(key);
        }
    });
}

// 生成Excel模板
function generateExcelTemplate() {
    // 生成2026年1月至12月的月度数据
    const templateData = [];
    for (let month = 1; month <= 12; month++) {
        templateData.push({
            '月度': `2026-${String(month).padStart(2, '0')}`,
            '行政服务满意度评分': '',
            '工单首次响应时长(小时)': '',
            '住宿费用(MXN)': '',
            '用车费用(MXN)': '',
            '餐饮费用(MXN)': '',
            '办公费用(MXN)': '',
            '供应商SLA合格率(%)': '',
            '月度入境墨西哥员工总人数': '',
            '成功入境墨西哥员工人数': ''
        });
    }
    
    const ws = XLSX.utils.json_to_sheet(templateData);
    
    // 设置列宽和格式
    const wscols = [
        {wch: 12},  // 月度
        {wch: 20},  // 行政服务满意度评分
        {wch: 20},  // 工单首次响应时长(小时)
        {wch: 18},  // 住宿费用(MXN)
        {wch: 18},  // 用车费用(MXN)
        {wch: 18},  // 餐饮费用(MXN)
        {wch: 18},  // 办公费用(MXN)
        {wch: 20},  // 供应商SLA合格率(%)
        {wch: 22},  // 月度入境墨西哥员工总人数
        {wch: 20}   // 成功入境墨西哥员工人数
    ];
    ws['!cols'] = wscols;
    
    // 设置单元格格式为文本格式
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; ++R) {
        for (let C = 0; C <= range.e.c; ++C) {
            const cell_address = {c:C, r:R};
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            if (!ws[cell_ref]) continue;
            ws[cell_ref].t = 's'; // 设置为字符串类型
        }
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '数据');
    
    XLSX.writeFile(wb, '行政数据模板.xlsx');
}

// 校验Excel列名
function validateExcelColumns(headers) {
    const requiredColumns = [
        '月度',
        '行政服务满意度评分',
        '工单首次响应时长(小时)',
        '住宿费用(MXN)',
        '用车费用(MXN)',
        '餐饮费用(MXN)',
        '办公费用(MXN)',
        '供应商SLA合格率(%)',
        '月度入境墨西哥员工总人数',
        '成功入境墨西哥员工人数'
    ];
    
    const missingColumns = requiredColumns.filter(col => !headers.includes(col));
    return missingColumns;
}

// 校验数据格式
function validateDataFormat(data) {
    const errors = [];
    
    data.forEach((row, index) => {
        const rowNum = index + 2; // 从第2行开始（第1行是表头）
        
        // 校验月度格式
        const monthRegex = /^\d{4}-\d{2}$/;
        if (!monthRegex.test(row['月度'])) {
            errors.push(`第 ${rowNum} 行：月度格式错误，应为 YYYY-MM`);
        }
        
        // 校验满意度评分（0-10）
        const satisfaction = parseFloat(row['行政服务满意度评分']);
        if (isNaN(satisfaction) || satisfaction < 0 || satisfaction > 10) {
            errors.push(`第 ${rowNum} 行：满意度评分应为 0-10 之间的数字`);
        }
        
        // 校验响应时长（正数）
        const responseTime = parseFloat(row['工单首次响应时长(小时)']);
        if (isNaN(responseTime) || responseTime < 0) {
            errors.push(`第 ${rowNum} 行：工单首次响应时长应为正数`);
        }
        
        // 校验费用（正数）
        const expenses = [
            row['住宿费用(MXN)'],
            row['用车费用(MXN)'],
            row['餐饮费用(MXN)'],
            row['办公费用(MXN)']
        ];
        
        expenses.forEach((expense, i) => {
            const expenseName = ['住宿费用', '用车费用', '餐饮费用', '办公费用'][i];
            const value = parseFloat(expense);
            if (isNaN(value) || value < 0) {
                errors.push(`第 ${rowNum} 行：${expenseName} 应为正数`);
            }
        });
        
        // 校验SLA合格率（0-100）
        const sla = parseFloat(row['供应商SLA合格率(%)']);
        if (isNaN(sla) || sla < 0 || sla > 100) {
            errors.push(`第 ${rowNum} 行：SLA合格率应为 0-100 之间的数字`);
        }
        
        // 校验员工人数（正整数）
        const totalEmployees = parseInt(row['月度入境墨西哥员工总人数']);
        const successEmployees = parseInt(row['成功入境墨西哥员工人数']);
        
        if (isNaN(totalEmployees) || totalEmployees < 1) {
            errors.push(`第 ${rowNum} 行：总人数应为正整数`);
        }
        
        if (isNaN(successEmployees) || successEmployees < 1) {
            errors.push(`第 ${rowNum} 行：成功入境人数应为正整数`);
        }
        
        if (successEmployees > totalEmployees) {
            errors.push(`第 ${rowNum} 行：成功入境人数不能超过总人数`);
        }
    });
    
    return errors;
}

// 检查重复月度数据
function checkDuplicateMonths(newData, existingData) {
    const existingMonths = existingData.map(item => item.month);
    const duplicateMonths = newData
        .map(item => item.month)
        .filter(month => existingMonths.includes(month));
    
    return duplicateMonths;
}

// 处理Excel数据
function processExcelData(data) {
    return data.map(row => ({
        month: row['月度'],
        satisfactionScore: parseFloat(row['行政服务满意度评分']),
        responseTime: parseFloat(row['工单首次响应时长(小时)']),
        accommodationExpenses: parseFloat(row['住宿费用(MXN)']),
        vehicleExpenses: parseFloat(row['用车费用(MXN)']),
        cateringExpenses: parseFloat(row['餐饮费用(MXN)']),
        officeExpenses: parseFloat(row['办公费用(MXN)']),
        slaCompliance: parseFloat(row['供应商SLA合格率(%)']),
        totalEmployees: parseInt(row['月度入境墨西哥员工总人数']),
        successEmployees: parseInt(row['成功入境墨西哥员工人数'])
    }));
}

// 保存数据到localStorage
function saveDataToLocalStorage(data) {
    localStorage.setItem('adminData', JSON.stringify(data));
}

// 从localStorage获取数据
function getDataFromLocalStorage() {
    const data = localStorage.getItem('adminData');
    return data ? JSON.parse(data) : [];
}

// 格式化数字
function formatNumber(num, decimals = 2) {
    if (num === null || num === undefined) return '-';
    return num.toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// 格式化货币
function formatCurrency(num) {
    if (num === null || num === undefined) return '-';
    return 'MXN ' + num.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// 格式化百分比
function formatPercentage(num) {
    if (num === null || num === undefined) return '-';
    return num.toFixed(1) + '%';
}

// 生成下载文件名
function generateDownloadFileName(format) {
    const today = new Date();
    const dateStr = today.getFullYear() + 
        String(today.getMonth() + 1).padStart(2, '0') + 
        String(today.getDate()).padStart(2, '0');
    return `墨西哥工厂行政服务数据_${dateStr}.${format}`;
}

// 下载Excel文件
function downloadExcel(data, filename) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Datos');
    XLSX.writeFile(wb, filename);
}

// 下载CSV文件
function downloadCSV(data, filename) {
    const csvContent = 'data:text/csv;charset=utf-8,' + 
        data.map(e => Object.values(e).join(',')).join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// 显示模态框
function showModal(title, message, onConfirm = null, onCancel = null) {
    const modal = document.getElementById('modal');
    const modalTitle = document.getElementById('modal-title');
    const modalMessage = document.getElementById('modal-message');
    const modalConfirm = document.getElementById('modal-confirm');
    const modalCancel = document.getElementById('modal-cancel');
    
    modalTitle.textContent = title;
    modalMessage.textContent = message;
    modal.classList.remove('hidden');
    
    modalConfirm.onclick = () => {
        modal.classList.add('hidden');
        if (onConfirm) onConfirm();
    };
    
    modalCancel.onclick = () => {
        modal.classList.add('hidden');
        if (onCancel) onCancel();
    };
    
    // 点击关闭按钮
    document.querySelector('.modal-close').onclick = () => {
        modal.classList.add('hidden');
    };
    
    // 点击模态框外部关闭
    modal.onclick = (e) => {
        if (e.target === modal) {
            modal.classList.add('hidden');
        }
    };
}

// 显示提示消息
function showAlert(message) {
    showModal(t('modal_title'), message);
}