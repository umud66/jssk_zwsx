(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.ResidencyRules = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';
    const PROVINCE_SOURCE = 'https://www.jszzb.gov.cn/tzgg/art/2025/art_0b70c8f3fb4f2c5d852180928e946634.html';
    const PROVINCE_SOURCES = {
        2026: PROVINCE_SOURCE,
        2027: 'https://www.jszzb.gov.cn/2027ndkslygwyzt/zkgg/art/2026/art_f9d91bae1e764231843f0be62ffadffa.html'
    };
    // 县区名称对应当前2026职位数据；用于避免仅凭“鼓楼区”等重名地区误匹配。
    const AREAS = {
        '南京市': ['玄武区','秦淮区','建邺区','鼓楼区','雨花台区','栖霞区','浦口区','六合区','江宁区','溧水区','高淳区','江北新区'],
        '无锡市': ['梁溪区','锡山区','惠山区','滨湖区','新吴区','江阴市','宜兴市'],
        '徐州市': ['鼓楼区','云龙区','贾汪区','泉山区','铜山区','丰县','沛县','睢宁县','新沂市','邳州市'],
        '常州市': ['天宁区','钟楼区','新北区','武进区','金坛区','溧阳市','常州经开区'],
        '苏州市': ['虎丘区','吴中区','相城区','姑苏区','吴江区','常熟市','张家港市','昆山市','太仓市','苏州工业园区'],
        '南通市': ['崇川区','通州区','海门区','如东县','启东市','如皋市','海安市'],
        '连云港市': ['连云区','海州区','赣榆区','东海县','灌云县','灌南县','连云港徐圩新区'],
        '淮安市': ['淮安区','淮阴区','清江浦区','洪泽区','涟水县','盱眙县','金湖县'],
        '盐城市': ['亭湖区','盐都区','大丰区','响水县','滨海县','阜宁县','射阳县','建湖县','东台市'],
        '扬州市': ['广陵区','邗江区','江都区','宝应县','仪征市','高邮市'],
        '镇江市': ['京口区','润州区','丹徒区','丹阳市','扬中市','句容市'],
        '泰州市': ['海陵区','高港区','姜堰区','兴化市','靖江市','泰兴市'],
        '宿迁市': ['宿城区','宿豫区','沭阳县','泗阳县','泗洪县','湖滨新区']
    };
    const POLICIES = {
        '无锡市': { socialArea: '无锡市', source: 'https://zzb.wuxi.gov.cn/doc/2025/11/01/4674865.shtml' },
        '苏州市': { socialArea: '苏州市', source: 'https://www.suzhou.gov.cn/szsrmzf/gwyjsydwklxx/202510/948caf3762bd4228b65e7368463acce1.shtml' },
        '连云港市': { socialArea: '江苏省', source: 'https://www.lyg.gov.cn/zglygzfmhwz/gbxbjgwysydwzlxx/content/4322124f-ef40-40ad-b858-7c0ff1565d9d.shtml' },
        '盐城市': { unrestricted: true, source: 'https://www.hsyf.gov.cn/p/9344.html' },
        '淮安市': { unrestricted: true, source: 'https://zzb.huaian.gov.cn/col/13489_621588/art/17592480/1761882935398Ng0a54vu.html' },
        '泰州市': { unrestricted: true, source: 'https://www.taixing.gov.cn/zwgk/xxgk/gwyzk/art/2026/art_254c7adca7d5460bb2cbd6eedc987519.html' }
    };
    const text = value => typeof value === 'string' ? value.trim() : '';
    const normalize = value => text(value).replace(/[\s/、，,·-]/g, '').replace(/（/g, '(').replace(/）/g, ')');
    const short = value => value.replace(/[省市县区]$/, '');
    const isGraduate = education => ['yes', 'regular'].includes(education.graduate);
    const isPostgraduate = education => ['master', 'doctor'].includes(education.level) && education.degree === 'yes';
    function parsePlace(value) {
        const raw = normalize(value);
        const foreign = raw.match(/^(.+?省|.+?自治区|北京市|上海市|天津市|重庆市)/)?.[1];
        if (foreign && foreign !== '江苏省') return { province: foreign, city: '', district: '', raw };
        let city = Object.keys(AREAS).find(name => raw.includes(name) || raw.includes(short(name))) || '';
        const candidates = [];
        for (const [parent, districts] of Object.entries(AREAS)) {
            if (city && parent !== city) continue;
            for (const district of districts) {
                const alias = short(district);
                if (raw.includes(district) || (alias.length >= 2 && alias !== short(parent) && raw.includes(alias))) candidates.push({ city: parent, district });
            }
        }
        let district = '';
        if (candidates.length === 1) { city = candidates[0].city; district = candidates[0].district; }
        return { province: city || /江苏/.test(raw) ? '江苏省' : '', city, district, raw };
    }
    function jobPlace(job) {
        const region = parsePlace(job.rn);
        const city = AREAS[job.c] ? job.c : region.city;
        return { province: '江苏省', city, district: region.district, raw: text(job.rn) };
    }
    function requiredPlace(area, job) {
        const context = jobPlace(job), raw = normalize(area);
        if (/本省|我省/.test(raw)) return { province: '江苏省', city: '', district: '' };
        if (/本县|本区/.test(raw)) return context.district ? context : null;
        if (/本市|我市/.test(raw)) {
            if (context.district && context.district.endsWith('市')) return context;
            return context.city ? { province: '江苏省', city: context.city, district: '' } : null;
        }
        const parsed = parsePlace(raw.replace(/^(?:面向|仅限|限|须为|应为|要求|但)+/, '').replace(/\([^)]*\)/g, ''));
        return parsed.province ? parsed : null;
    }
    function matchesPlace(value, required) {
        if (!text(value) || !required) return false;
        const place = parsePlace(value);
        return place.province === required.province && (!required.city || place.city === required.city) && (!required.district || place.district === required.district);
    }
    function splitClauses(value) {
        let depth = 0, part = '';
        const result = [];
        for (const char of text(value)) {
            if (char === '(' || char === '（') depth++;
            if (char === ')' || char === '）') depth = Math.max(0, depth - 1);
            if (!depth && /[，,；;。\n]/.test(char)) { if (part.trim()) result.push(part.trim()); part = ''; }
            else part += char;
        }
        if (part.trim()) result.push(part.trim());
        return result;
    }
    function isLocationClause(clause) {
        return /户籍|生源/.test(clause) && !/退役|军士|基层项目|残疾|优秀村|公安院校|配偶|父母/.test(clause);
    }
    function explicitRestrictions(job, education) {
        const restrictions = [];
        for (const clause of splitClauses(job.ot)) {
            if (!/户籍|生源/.test(clause) || /退役|军士|基层项目|配偶|父母/.test(clause)) continue;
            const waiverStart = clause.search(/(?:研究生|硕士|博士).{0,35}(?:不受|不限|不限制)/);
            const limitClause = waiverStart >= 0 ? clause.slice(0, waiverStart) : clause;
            if (!/户籍|生源/.test(limitClause) || /不受|不限|不限制|无.{0,6}限制/.test(limitClause)) continue;
            const graduatePart = isGraduate(education) ? limitClause.match(/应届(?:生|毕业生)?(?:须为|限|应为|为)([^)）]+)/)?.[1] : '';
            const effective = graduatePart || limitClause.replace(/[（(]应届[^)）]*[)）]/g, '');
            const index = effective.search(/户籍|生源/);
            if (index < 0) continue;
            const prefix = effective.slice(0, index).replace(/\([^)]*\)|（[^）]*）/g, '');
            const area = requiredPlace(prefix, job);
            const household = /户籍/.test(effective), origin = /生源/.test(effective);
            restrictions.push({ area, type: household && origin ? (/且|同时/.test(effective) ? 'both' : 'either') : household ? 'household' : 'origin', raw: clause });
        }
        return restrictions;
    }
    function result(status, reason, source = '') { return { status, reason, source }; }
    function checkRestrictions(restrictions, job, education, profile) {
        const raw = text(job.ot);
        const waiver = raw.match(/(研究生|硕士及以上|博士|硕士)[^。；;]{0,25}(?:不受|不限|不限制)((?:本市|本县|本区|生源地|生源|户籍|地域|、|或|和|及|与|的|此|限制|\s){1,25})/);
        const correctLevel = waiver && (waiver[1] === '博士' ? education.level === 'doctor' : waiver[1] === '硕士' ? education.level === 'master' : ['master', 'doctor'].includes(education.level));
        const pgWaiver = isPostgraduate(education) && correctLevel;
        const both = pgWaiver && (/地域/.test(waiver[2]) || (/此限制/.test(waiver[2]) && restrictions.length === 1));
        const waiveHousehold = pgWaiver && (both || /户籍/.test(waiver[2]));
        const waiveOrigin = pgWaiver && (both || /生源/.test(waiver[2]));
        for (const restriction of restrictions) {
            if (!restriction.area) return result(profile.household || profile.origin ? 'review' : 'mismatch', '岗位地域指代需核对；请填写完整户籍与生源地区');
            const household = waiveHousehold || matchesPlace(profile.household, restriction.area), origin = waiveOrigin || matchesPlace(profile.origin, restriction.area);
            const matches = restriction.type === 'either' ? household || origin : restriction.type === 'both' ? household && origin : restriction.type === 'household' ? household : origin;
            if (!matches) return result('mismatch', `未满足岗位明确地域要求：${restriction.raw}`);
        }
        if (waiveHousehold || waiveOrigin) return result('match', `以该条学历专业及相应学位报考，适用岗位原文的${[waiveHousehold ? '户籍' : '', waiveOrigin ? '生源' : ''].filter(Boolean).join(' / ')}豁免`);
        return result('match', `已核对岗位明确的${restrictions.map(item => item.type === 'origin' ? '生源' : item.type === 'household' ? '户籍' : '户籍 / 生源').join('、')}要求`);
    }
    function match(job, education, profile, year = 2026) {
        const restrictions = explicitRestrictions(job, education);
        if (restrictions.length) return checkRestrictions(restrictions, job, education, profile);
        const raw = text(job.ot);
        const unlimited = splitClauses(raw).some(clause => /户籍|生源/.test(clause) && /不限|不受|不限制/.test(clause) && !/研究生|硕士|博士|本科|应届/.test(clause));
        if (unlimited) return result('match', '岗位原文明确不限制户籍 / 生源');
        const context = jobPlace(job), policy = POLICIES[context.city], code = Number(job.pc);
        const postgraduate = isPostgraduate(education);
        const provinceSource = PROVINCE_SOURCES[year];
        if (!provinceSource) return result('review', `${year}年度户籍 / 生源政策尚未纳入，请核对当年度简章`);
        // 定向应届职位的江苏生源要求优先于一般应届本科豁免。
        if ((code >= 60 && code <= 69) || new RegExp(`面向${year}年普通高校应届毕业生`).test(raw)) {
            if (postgraduate) return result('match', '以该条研究生专业及相应学位报考，适用定向应届职位的生源豁免', provinceSource);
            if (matchesPlace(profile.origin, { province: '江苏省' })) return result('match', '定向应届职位：已核对江苏生源', provinceSource);
            return result('mismatch', '该定向应届职位须为江苏生源；未填生源不能满足要求', provinceSource);
        }
        if (job.c === '省级' || (year === 2026 && code >= 50 && code <= 59)) return result('match', '', provinceSource);
        if (job.c === '监狱戒毒') return result('review', '此岗位适用司法行政专项简章，户籍 / 生源规则需另行核对', provinceSource);
        if (!context.city) return result('match', '');
        // 地方简章按年度生效，未核对的新年度不能套用旧年度的准入或豁免。
        if (year !== 2026) return result('review', `${year}年度${context.city}户籍 / 生源通用政策待核对；已检查岗位原文地域限制，未套用2026年度规则`, provinceSource);
        if (!policy) return result('review', `${context.city}的通用户籍政策尚未纳入，不能据此确认地域资格`, PROVINCE_SOURCE);
        if (policy.unrestricted) return result('match', `${context.city}一般职位不限制户籍 / 生源；已另行检查岗位特殊规定`, policy.source);
        const regular = education.graduate === 'regular';
        if (postgraduate && (!isGraduate(education) || regular)) return result('match', `以该条研究生专业及相应学位报考，适用${context.city}通用户籍 / 生源豁免`, policy.source);
        const hasCurrentGraduate = profile.educations.some(item => ['master', 'doctor'].includes(item.level) && ['yes', 'regular'].includes(item.graduate));
        if (education.level === 'bachelor' && regular && !hasCurrentGraduate) return result('match', `以该条普通高校2026应届本科学历专业报考，适用${context.city}一般职位生源豁免`, policy.source);
        if (context.city === '苏州市' && education.level === 'bachelor' && profile.educations.some(item => ['master', 'doctor'].includes(item.level) && item.graduate === 'regular')) {
            if (matchesPlace(profile.household, { province: '江苏省', city: '苏州市' }) || matchesPlace(profile.origin, { province: '江苏省', city: '苏州市' })) return result('match', '应届研究生以本科专业报考：已核对苏州户籍或苏州生源', policy.source);
            return result('mismatch', '应届研究生以本科专业报考仍须苏州户籍或苏州生源，不能借用研究生豁免', policy.source);
        }
        const required = requiredPlace(policy.socialArea, job);
        if (matchesPlace(profile.household, required)) return result('match', `按该条学历专业报考，已核对${policy.socialArea}户籍`, policy.source);
        return result('mismatch', `此报考学历专业路径要求${policy.socialArea}户籍，未满足；不能借用其他学历的地域豁免`, policy.source);
    }
    const suggestions = () => ['外省', ...Object.keys(AREAS).map(city => `江苏省${city}`)];
    return Object.freeze({ match, parsePlace, matchesPlace, isLocationClause, suggestions, policies: POLICIES });
}));
