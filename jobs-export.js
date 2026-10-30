(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.JobExport = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';
    const columns = [
        ['城市 / 系统', 'c', 16], ['地区', 'rn', 18], ['单位名称', 'un', 32],
        ['职位名称', 'pn', 38], ['招考人数', 'rh', 12], ['学历要求', 'ed', 18],
        ['专业要求', 'mj', 58], ['其他要求', 'ot', 70], ['职位简介', 'pd', 54],
        ['考试类别', 'et', 12], ['隶属关系', 'r', 12], ['地区代码', 'rc', 14],
        ['单位代码', 'uc', 12], ['职位代码', 'pc', 12], ['开考比例', 'kr', 12]
    ];
    const profileColumns = ['初筛结果', '匹配学历专业', '专业匹配说明', '户籍 / 生源判断', '其他核对事项', '政策依据'];
    const text = value => value == null ? '' : String(value);
    const unique = values => [...new Set(values.filter(Boolean))];
    const numberOrText = value => text(value).trim() && Number.isFinite(Number(value)) ? Number(value) : text(value);

    function matchCells(result) {
        const matches = result.matches || [];
        const residence = unique(matches.map(item => item.residency?.reason));
        const otherReviews = unique((result.reviews || []).filter(reason => !residence.includes(reason)));
        const label = result.status === 'match' ? '资料初筛匹配' : otherReviews.length ? '待核对' : '户籍待核实';
        const groups = new Map();
        matches.filter(item => item.residency?.reason).forEach(item => {
            const reason = item.residency.reason;
            if (!groups.has(reason)) groups.set(reason, []);
            groups.get(reason).push(item.index + 1);
        });
        return [label,
            matches.map(item => `学历 ${item.index + 1}：${item.level} · ${item.major}`).join('；'),
            unique(matches.map(item => item.reason)).join('；'),
            [...groups].map(([reason, indices]) => `${matches.length > 1 ? `学历 ${indices.join('、')}：` : ''}${reason}`).join('\n'),
            otherReviews.join('；'), unique(matches.map(item => item.residency?.source)).join('\n')];
    }

    function createWorkbook(XLSX, jobs, options) {
        if (!Array.isArray(jobs) || !jobs.length) throw new Error('没有可导出的职位');
        const { year, profileEnabled = false, matches = new Map(), filters = [],
            scope = '全部职位', sort = '默认排序', includeReview = true, createdAt = new Date() } = options;
        const headers = ['序号', '招录年度', ...columns.map(item => item[0]), ...(profileEnabled ? profileColumns : [])];
        const rows = jobs.map((job, index) => {
            const row = [index + 1, year, ...columns.map(([, key]) => ['rh', 'kr'].includes(key) ? numberOrText(job[key]) : text(job[key]))];
            if (profileEnabled) {
                if (!matches.has(job)) throw new Error('缺少职位的初筛结果，请重新筛选后导出');
                row.push(...matchCells(matches.get(job)));
            }
            return row;
        });
        const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        sheet['!cols'] = [{ wch: 8 }, { wch: 12 }, ...columns.map(item => ({ wch: item[2] })), ...(profileEnabled ? [18, 48, 50, 70, 60, 65].map(wch => ({ wch })) : [])];
        sheet['!autofilter'] = { ref: sheet['!ref'] };
        const summary = [
            ['项目', '内容'], ['招录年度', year], ['导出时间', createdAt.toLocaleString('zh-CN', { hour12: false })],
            ['职位范围', scope], ['导出职位数', jobs.length],
            ['招考人数合计', jobs.reduce((sum, job) => sum + (Number(job.rh) || 0), 0)],
            ['排序', sort], ['个人资料初筛', profileEnabled ? '已启用' : '未启用'],
            ...(profileEnabled ? [['待核对职位', includeReview ? '包含' : '不包含']] : []),
            ...(filters.length ? filters.map(filter => ['筛选条件', text(filter)]) : [['筛选条件', '无额外条件']])
        ];
        const conditions = XLSX.utils.aoa_to_sheet(summary);
        conditions['!cols'] = [{ wch: 22 }, { wch: 90 }];
        const workbook = XLSX.utils.book_new();
        workbook.Props = { Title: `江苏省${year}年度公务员筛选职位`, Subject: scope, CreatedDate: createdAt };
        XLSX.utils.book_append_sheet(workbook, sheet, '岗位清单');
        XLSX.utils.book_append_sheet(workbook, conditions, '筛选条件');
        return workbook;
    }

    function filename(year, count, date = new Date()) {
        const pad = value => String(value).padStart(2, '0');
        const stamp = `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
        return `江苏省${year}年_筛选职位_${count}个_${stamp}.xlsx`;
    }

    return Object.freeze({ createWorkbook, filename });
}));
