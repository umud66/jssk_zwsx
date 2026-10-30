(function (root, factory) {
    const api = factory(typeof module === 'object' && module.exports ? require('./major-catalog-2027.js') : root.MajorCatalog2027);
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.MajorCatalog = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, function (data) {
    'use strict';
    const normalize = value => typeof value === 'string' ? value.replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')').toLowerCase() : '';
    const group = level => ({ master: 'postgraduate', doctor: 'postgraduate', bachelor: 'bachelor', college: 'college' })[level];
    const indexes = new Map(), lists = new Map();
    const available = year => !!data && data.year === Number(year);
    const generic = name => /各类|分语文、数学|等具体学科/.test(name);
    const categoryNames = new Set((data?.categories || []).map(category => normalize(category.name)));
    const isPlaceholder = (year, name) => available(year) && (categoryNames.has(normalize(name)) || generic(name));
    for (const level of ['postgraduate', 'bachelor', 'college']) {
        const index = new Map();
        const add = (name, category) => {
            if (!name || generic(name)) return;
            const key = normalize(name);
            if (!index.has(key)) index.set(key, { name, categories: [] });
            const entry = index.get(key);
            if (!entry.categories.includes(category)) entry.categories.push(category);
        };
        for (const category of data?.categories || []) {
            for (const name of category[level]) {
                add(name, category.name);
                const annotation = name.match(/^([^（(]+)[（(](.*)[）)]$/);
                if (!annotation || !/^(?:含[:：]?|分语族|可注明)/.test(annotation[2])) continue;
                add(annotation[1], category.name);
                if (/^含[:：]?/.test(annotation[2])) {
                    annotation[2].replace(/^含[:：]?/, '').split(/[、，,]/).forEach(item => add(item, category.name));
                }
            }
        }
        indexes.set(level, index);
        lists.set(level, [...index.values()].map(entry => entry.name).sort((a, b) => a.localeCompare(b, 'zh-CN')));
    }
    function lookup(year, level, major) {
        if (!available(year)) return null;
        const entry = indexes.get(group(level))?.get(normalize(major));
        return entry ? { name: entry.name, categories: [...entry.categories] } : null;
    }
    function names(year, level) {
        return available(year) ? [...(lists.get(group(level)) || [])] : [];
    }
    function distance(a, b) {
        let previous = Array.from({ length: b.length + 1 }, (_, i) => i);
        for (let i = 1; i <= a.length; i++) {
            const current = [i];
            for (let j = 1; j <= b.length; j++) current[j] = Math.min(current[j - 1] + 1, previous[j] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
            previous = current;
        }
        return previous[b.length];
    }
    function suggestions(year, level, query, limit = 6) {
        const value = normalize(query);
        if (!value || lookup(year, level, query)) return [];
        const maximum = value.length < 4 ? 1 : 2;
        return names(year, level).map(name => {
            const key = normalize(name);
            const substring = value.length >= 2 && key.includes(value);
            const score = substring ? 0 : Math.abs(key.length - value.length) <= maximum ? distance(value, key) : Infinity;
            return { name, score };
        }).filter(item => item.score <= maximum)
            .sort((a, b) => a.score - b.score || a.name.length - b.name.length || a.name.localeCompare(b.name, 'zh-CN'))
            .slice(0, Math.min(10, Math.max(0, limit))).map(item => item.name);
    }
    return Object.freeze({ available, lookup, names, suggestions, isPlaceholder });
}));
