(function (root, factory) {
    const api = factory(typeof module === 'object' && module.exports ? require('./residency-rules.js') : root.ResidencyRules);
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.ProfileMatcher = api;
}(typeof globalThis !== 'undefined' ? globalThis : this, function (residencyRules) {
    'use strict';
    const LEVELS = Object.freeze({ highschool: 1, college: 2, bachelor: 3, master: 4, doctor: 5 });
    const LABELS = Object.freeze({ highschool: '高中 / 中专', college: '大专', bachelor: '本科', master: '硕士研究生', doctor: '博士研究生' });
    const clean = value => typeof value === 'string' ? value.trim().slice(0, 160) : '';
    const text = value => typeof value === 'string' ? value.trim() : '';
    const normalize = value => text(value).replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')').toLowerCase();
    const allowed = (value, values) => values.includes(value) ? value : '';
    const chineseNumber = value => /^\d+(\.\d+)?$/.test(value) ? Number(value) : ({ 一: 1, 二: 2, 两: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 }[value] ?? null);
    function normalizeProfile(value) {
        const source = value && typeof value === 'object' ? value : {};
        const years = typeof source.workYears === 'number' || typeof source.workYears === 'string' ? String(source.workYears).trim() : '';
        return {
            version: 1,
            enabled: source.enabled === true,
            includeReview: source.includeReview !== false,
            gender: allowed(source.gender, ['male', 'female']),
            politics: allowed(source.politics, ['party', 'probationary', 'league', 'masses', 'other']),
            examType: allowed(source.examType, ['A', 'B', 'C']),
            household: clean(source.household),
            origin: clean(source.origin),
            workYears: years !== '' && Number.isFinite(Number(years)) && Number(years) >= 0 && Number(years) <= 70 ? String(Number(years)) : '',
            credentialKeywords: text(source.credentialKeywords).slice(0, 4000),
            eligibilityKeywords: text(source.eligibilityKeywords).slice(0, 4000),
            strictCredentials: true,
            educations: Array.isArray(source.educations) ? source.educations.filter(item => item && LEVELS[item.level] && clean(item.major)).map(item => ({
                level: item.level, major: clean(item.major), categories: clean(item.categories),
                degree: allowed(item.degree, ['yes', 'no']), graduate: allowed(item.graduate, ['yes', 'no', 'regular'])
            })) : []
        };
    }
    function educationMatch(requirement, level) {
        const text = normalize(requirement), rank = LEVELS[level];
        if (!rank) return 'mismatch';
        if (!text || /^(不限|不限学历|学历不限)$/.test(text)) return text ? 'match' : 'review';
        let minimum;
        if (/博士/.test(text)) minimum = 5;
        else if (/硕士/.test(text)) minimum = 4;
        else if (/^研究生(及以上|以上)?$/.test(text)) return rank >= 4 ? 'match' : 'mismatch';
        else if (/研究生/.test(text)) return 'review';
        else if (/本科/.test(text)) minimum = 3;
        else if (/大专|专科/.test(text)) minimum = 2;
        else if (/高中|中专/.test(text)) minimum = 1;
        else return 'review';
        // 只自动处理职位表中明确的单层次要求；混合条件需人工核对。
        const simple = /^(博士研究生|博士|硕士研究生|硕士|本科|大学本科|大专|大学专科|专科|高中\(中专\)|高中|中专)(及以上|以上)?$/.test(text);
        if (!simple) return 'review';
        return (/以上/.test(text) ? rank >= minimum : rank === minimum) ? 'match' : 'mismatch';
    }
    function splitMajors(text) {
        const parts = [];
        let depth = 0, part = '';
        for (const char of normalize(text)) {
            if (char === '(') depth++;
            if (char === ')') depth = Math.max(0, depth - 1);
            if (!depth && /[,，、;；\n]/.test(char)) { if (part) parts.push(part); part = ''; }
            else part += char;
        }
        if (part) parts.push(part);
        return parts;
    }
    function majorMatch(requirement, education) {
        const text = normalize(requirement), major = normalize(education.major);
        if (!text) return { status: 'review', reason: '岗位未注明专业，需核对' };
        if (/^(不限|专业不限|不限专业|不限制专业)$/.test(text)) return { status: 'match', reason: '岗位不限专业' };
        if (/^(不限|专业不限|不限专业|不限制专业)\(/.test(text)) return { status: 'review', reason: '不限专业但有附加条件，请核对原文' };
        const parts = splitMajors(text), categories = splitMajors(education.categories);
        if (parts.includes(major)) return { status: 'match', reason: `专业名称匹配：${education.major}` };
        const category = categories.find(value => parts.includes(value));
        if (category) return { status: 'match', reason: `按已填写专业类别匹配：${category}` };
        const possibleCategory = parts.find(value => value.endsWith('类') && (major.startsWith(value.slice(0, -1)) || (value === '法律类' && /^(法学|法律)(\(.*\))?$/.test(major))));
        if (possibleCategory) return { status: 'review', reason: `可能属于${possibleCategory}，请确认该学历的专业类别` };
        if (parts.some(value => !value.endsWith('类') && (value.includes(major) || major.includes(value)))) return { status: 'review', reason: '专业名称相近，需核对岗位要求的完整名称' };
        return { status: 'mismatch', reason: '专业不匹配' };
    }
    function globalConditions(job, profile) {
        const requirement = text(job.ot), reviews = [];
        if (profile.examType && text(job.et).replace(/\s|类/g, '').toUpperCase() !== profile.examType) return { status: 'mismatch', reviews: ['考试类型不匹配'] };
        const strict = requirement.replace(/(?:适合|适宜|建议)[男女]性/g, '');
        const male = strict.includes('男性'), female = strict.includes('女性');
        if (male !== female) {
            if (!profile.gender) return { status: 'mismatch', reviews: ['未填写岗位要求的性别信息'] };
            else if ((male && profile.gender !== 'male') || (female && profile.gender !== 'female')) return { status: 'mismatch', reviews };
        }
        for (const clause of splitOutside(requirement, /[，,；;。\n]/).filter(value => /中共党员|共青团员|民主党派/.test(value))) {
            const matches = (profile.politics === 'party' && /中共党员/.test(clause))
                || (profile.politics === 'probationary' && /中共党员/.test(clause) && /含预备|包括预备/.test(clause))
                || (profile.politics === 'league' && /共青团员/.test(clause));
            if (!matches) return { status: 'mismatch', reviews: ['未填写或不具备岗位要求的政治面貌'] };
        }
        const years = requirement.match(/(?:具有|具备|有)([一二两三四五六七八九十\d.]+)年(?:及)?以上基层工作经历/);
        if (years && chineseNumber(years[1]) !== null) {
            const minimum = chineseNumber(years[1]);
            if (Number(profile.workYears || 0) < minimum) return { status: 'mismatch', reviews: [`基层工作经历不足${minimum}年，未填写按0年处理`] };
        }
        return { status: reviews.length ? 'review' : 'match', reviews };
    }
    function splitOutside(value, separators) {
        const parts = [];
        let depth = 0, part = '';
        for (const char of text(value)) {
            if (char === '(' || char === '（') depth++;
            if (char === ')' || char === '）') depth = Math.max(0, depth - 1);
            if (!depth && separators.test(char)) { if (part.trim()) parts.push(part.trim()); part = ''; }
            else part += char;
        }
        if (part.trim()) parts.push(part.trim());
        return parts;
    }
    function normalizeEligibility(value, job) {
        return normalize(value).replace(/^(面向|仅限|限|招录对象为)/, '').replace(/我省|江苏(?!省)/g, '江苏省')
            .replace(/服现役/g, '服役').replace(/五年/g, '5年')
            .replace(/户籍\(生源\)|户籍、生源/g, '户籍或生源')
            .replace(/本县|本区/g, text(job.rn || job.c)).replace(/[，,；;。]/g, '');
    }
    function eligibilityConditions(job, profile, education) {
        const entries = splitOutside(profile.eligibilityKeywords, /[；;\n]/).filter(value => !/不是|没有|未参加|未服务|不具备|^无/.test(value)).map(value => normalizeEligibility(value, {}));
        const groups = [];
        let block = [];
        const flush = () => { if (block.length) groups.push(block.join('，')); block = []; };
        for (const clause of splitOutside(text(job.ot), /[，,；;。\n]/)) {
            // 纯地域条件在每条学历内判断，不能在这里先否决整个岗位。
            if (residencyRules.isLocationClause(clause)) { flush(); continue; }
            // 这些完整的并列身份是 OR，不能把其中的退役条件应用给基层项目人员。
            const common = /服务基层项目人员|(?:服现役|服役)满(?:五|5)年的?[^、，；;。]*?高校毕业生退役军人|符合相关条件并在我省安排工作的退役军士和义务兵/g;
            const branches = clause.match(common) || [];
            const remainder = clause.replace(/^(面向|仅限|限)/, '').replace(common, '').replace(/[、或和及\s]/g, '');
            if (branches.length && !remainder) { flush(); groups.push(branches); continue; }
            const continuation = /曾在|服役期间|具备.*(?:特战|狙击)|须提供.*部队|提供.*团级|生活能够自理|能正常履行|独立开展工作|具有.*服役经历/.test(clause);
            if (block.length && (continuation || /^或/.test(clause))) { block.push(clause); continue; }
            const identity = /退役军人|退役士兵|退役人员|退役军士|复员士官|基层项目人员|优秀村|残疾|公安院校.*毕业生|健将级.*运动员/.test(clause);
            const locality = !/不限|不限制|无.*限制/.test(clause) && /(?:限|须为|要求).*?(?:户籍|生源)|^[^，；。]{1,25}户籍[（(]?生源/.test(clause);
            if (block.length && locality) { block.push(clause); continue; }
            flush();
            if (identity || locality) block.push(clause);
        }
        flush();
        const has = required => {
            // 退役分支内的本省户籍/生源不能因别的并列身份或一般研究生豁免而丢失。
            if (/退役/.test(required) && /户籍|生源/.test(required)) {
                const location = ['yes', 'regular'].includes(education.graduate) ? profile.origin : profile.household;
                if (!residencyRules.matchesPlace(location, { province: '江苏省' })) return false;
            }
            const target = normalizeEligibility(required, job);
            const variants = target.includes('户籍或生源') ? [target, target.replace('户籍或生源', '户籍'), target.replace('户籍或生源', '生源')] : [target];
            return entries.some(entry => variants.some(value => value && entry.includes(value)));
        };
        for (const group of groups) {
            if (Array.isArray(group)) {
                if (!group.some(has)) return { status: 'mismatch', reviews: ['未声明岗位要求的服务基层项目或指定退役身份'] };
            } else {
                const alternatives = group.split(/[，；;]或/);
                const matches = alternatives.some(branch => has(branch) || (/证书/.test(branch) && certificateCondition(branch, ownCredentials(profile)).status === 'match'));
                if (!matches) return { status: 'mismatch', reviews: ['未声明岗位要求的完整定向身份或限定条件'] };
            }
        }
        return { status: 'match', reviews: [] };
    }
    // 学历列与其他要求都具有约束力；已知的学位冲突不能降级为待核对。
    function degreeConditions(other, education, profile) {
        const reviews = [];
        let secondary = false;
        for (const clause of splitOutside(other, /[，,；;。\n]/)) {
            if (/以研究生学历报考/.test(clause) && LEVELS[education.level] < 4) continue;
            if (/本科阶段|本科段|本科为|本科所学|本科学历.*专业|本科.*并.*学位/.test(clause)) {
                secondary = true;
                const bachelors = profile.educations.filter(item => item.level === 'bachelor');
                if (!bachelors.length || (/学位/.test(clause) && !bachelors.some(item => item.degree === 'yes'))) return { status: 'mismatch', reviews: ['未填写岗位要求的本科阶段学历或学位'] };
                reviews.push('还需核对本科阶段的专业与学位要求');
            } else if (!/^(并|且|同时)/.test(clause)) secondary = false;
            if (!/学位/.test(clause) || !/取得|具有|获得|具备|须|要求/.test(clause)) continue;
            if (secondary) {
                if (!profile.educations.some(item => item.level === 'bachelor' && item.degree === 'yes')) return { status: 'mismatch', reviews: ['未填写岗位要求的本科阶段学位'] };
                continue;
            }
            const named = clause.match(/(博士|硕士|学士)(?:及以上|以上)?学位/);
            if (named) {
                const required = { 博士: 5, 硕士: 4, 学士: 3 }[named[1]];
                if (/或/.test(clause) && (clause.match(/博士|硕士|学士/g) || []).length > 1) {
                    reviews.push('学位含并列条件，请核对原文');
                } else {
                    const minimum = new RegExp(`${named[1]}(?:及以上|以上)学位|${named[1]}学位(?:及以上|以上)`).test(clause);
                    if (minimum ? LEVELS[education.level] < required : LEVELS[education.level] !== required) return { status: 'mismatch', reviews: [`岗位要求${named[1]}学位，与该条${LABELS[education.level]}资料不符`] };
                }
            }
            if (education.degree !== 'yes') return { status: 'mismatch', reviews: ['未填写或未取得岗位要求的学位'] };
            if (/法学学士|工学学士|理学学士|管理学学士|经济学学士/.test(clause)) reviews.push('指定学位门类需核对');
        }
        return { status: reviews.length ? 'review' : 'match', reviews: [...new Set(reviews)] };
    }
    function ownCredentials(profile, publications = false) {
        return splitOutside(profile.credentialKeywords, publications ? /[；;\n]/ : /[，,；;\n]/).filter(value => !/^(没有|暂无|尚未|未取得|未通过|未发表|未持有|无证|无论文|无成果)|没有论文|无论文|未取得|未通过|未发表/.test(value)).map(value => ({ original: value, key: normalize(value) }));
    }
    function legalGrade(value) {
        return normalize(value).match(/(?:法律(?:职业|执业)资格(?:证书|证)?|法考)\(?([abc])(?:类|证|\)|$)/)?.[1] || '';
    }
    function paperCondition(clause, entries) {
        const need = normalize(clause);
        const cssci = /cssci|中文社会科学引文索引/.test(need);
        const sci = !cssci && /(^|[^a-z])sci([^a-z]|$)/.test(need);
        const count = clause.match(/([\d一二两三四五六七八九十]+)篇/);
        const minimum = count ? chineseNumber(count[1]) : null;
        const label = `论文${cssci ? '（CSSCI）' : sci ? '（SCI）' : ''}${minimum ? `至少${minimum}篇` : ''}`;
        const related = entries.filter(entry => /论文|cssci|中文社会科学引文索引|(^|[^a-z])sci([^a-z]|$)/.test(entry.key));
        if (!related.length) return { status: 'missing', label, detail: `未填写成果：${label}` };
        if (/或/.test(need)) return { status: 'review', label, detail: `${label}含期刊或篇数并列条件，需核对原文` };
        const candidates = related.map(entry => {
            const missing = [];
            if (cssci && /(^|[^a-z])sci([^a-z]|$)/.test(entry.key) && !/cssci|中文社会科学引文索引/.test(entry.key)) return { status: 'mismatch', label, detail: '已填SCI论文与岗位要求的CSSCI类别不符' };
            if (sci && /cssci|中文社会科学引文索引/.test(entry.key) && !/(^|[^a-z])sci([^a-z]|$)/.test(entry.key)) return { status: 'mismatch', label, detail: '已填CSSCI论文与岗位要求的SCI类别不符' };
            if (cssci && !/cssci|中文社会科学引文索引/.test(entry.key)) missing.push('CSSCI期刊类别');
            if (sci && !/(^|[^a-z])sci([^a-z]|$)/.test(entry.key)) missing.push('SCI期刊类别');
            const ownCount = entry.original.match(/([\d一二两三四五六七八九十]+)篇/);
            if (minimum && ownCount && chineseNumber(ownCount[1]) !== null && chineseNumber(ownCount[1]) < minimum) return { status: related.length === 1 ? 'mismatch' : 'review', label, detail: `已填论文篇数不足${minimum}篇，需核对${label}` };
            if (minimum && !ownCount) missing.push(`至少${minimum}篇`);
            if (/第一作者/.test(need) && /第二作者|第三作者|通讯作者/.test(entry.key) && !/第一作者/.test(entry.key)) return { status: 'mismatch', label, detail: '已填作者身份不满足第一作者要求' };
            if (/第一作者/.test(need) && !/第一作者/.test(entry.key)) missing.push('第一作者身份');
            if (/读博以来|博士期间/.test(need) && /硕士期间|本科期间|读硕期间/.test(entry.key) && !/读博以来|博士期间|攻读博士期间/.test(entry.key)) return { status: 'mismatch', label, detail: '已填论文发表时间不满足读博以来的要求' };
            if (/读博以来|博士期间/.test(need) && !/读博以来|博士期间|攻读博士期间/.test(entry.key)) missing.push('读博以来的发表时间');
            if (/来源期刊/.test(need) && !/来源期刊/.test(entry.key)) missing.push('来源期刊范围');
            if (/不含扩展版/.test(need) && /扩展版/.test(entry.key) && !/不含扩展版|非扩展版/.test(entry.key)) return { status: 'mismatch', label, detail: '已填扩展版论文不满足岗位限定的来源期刊要求' };
            if (/不含[^)]*集刊/.test(need) && /集刊/.test(entry.key) && !/不含[^)]*集刊|非集刊/.test(entry.key)) return { status: 'mismatch', label, detail: '已填集刊论文不满足岗位限定的来源期刊要求' };
            if (sci && /[一二三四1234]区/.test(need)) missing.push('SCI分区要求');
            if (/毒理|药理|病理|药效|动物实验/.test(need)) missing.push('论文研究方向');
            return missing.length ? { status: 'review', label, detail: `${label}待确认：${missing.join('、')}` } : { status: 'match', label, detail: `已填成果信息包含${label}所列条件` };
        });
        return candidates.find(item => item.status === 'match') || candidates.find(item => item.status === 'review') || candidates[0];
    }
    function certificateCondition(clause, entries) {
        const need = normalize(clause);
        if (/法律(?:职业|执业)资格/.test(need)) {
            const grade = legalGrade(need), label = `国家法律职业资格证书${grade ? `（${grade.toUpperCase()}类）` : ''}`;
            const related = entries.filter(entry => /法律(?:职业|执业)资格|法考/.test(entry.key));
            if (!related.length) return { status: 'missing', label, detail: `未填写证书：${label}` };
            if (related.some(entry => !grade || legalGrade(entry.key) === grade)) return { status: 'match', label, detail: `已填写${label}` };
            if (related.some(entry => !legalGrade(entry.key))) return { status: 'review', label, detail: `${label}等级待确认；未写明${grade.toUpperCase()}类` };
            return { status: 'mismatch', label, detail: `已填证书等级与${label}要求不符` };
        }
        const target = need.replace(/^(?:并|且|须|应|要求|已经|已)?(?:取得|具有|获得|持有|通过|具备)/, '');
        const label = clause.replace(/^(?:并|且|须|应|要求|已经|已)?(?:取得|具有|获得|持有|通过|具备)/, '').trim();
        if (entries.some(entry => entry.key.includes(target))) return { status: 'match', label, detail: `已填写证书：${label}` };
        const related = entries.filter(entry => entry.key.length > 1 && (target.includes(entry.key) || entry.key.includes(target)));
        return related.length ? { status: 'review', label, detail: `证书名称相关，但等级或附加条件待确认：${label}` } : { status: 'missing', label, detail: `未确认具备：${label}` };
    }
    function credentialConditions(other, education, profile) {
        const entries = ownCredentials(profile), papers = ownCredentials(profile, true), groups = [];
        const isLaw = splitMajors(education.categories).includes('法律类') || /^(法学|法律)/.test(education.major);
        const isComputing = splitMajors(education.categories).includes('计算机类') || /^计算机/.test(education.major);
        for (const clause of splitOutside(other, /[，,；;。\n]/)) {
            if (/法律类专业/.test(clause) && !isLaw) continue;
            if (/非计算机类专业/.test(clause) && isComputing) continue;
            if (/退役|公安院校|健将级/.test(other) && (/^或/.test(clause) || /公安院校.*或.*证书/.test(clause))) continue;
            const publication = /论文|发表.*(?:CSSCI|SCI|期刊)/i.test(clause);
            const certificate = /证书|资格证|职业资格|专业技术资格|执业医师资格|注册会计师资格|翻译资格/.test(clause);
            const language = /大学英语|英语专业|雅思|托福/.test(clause) && /级|分|良好|优秀/.test(clause) && !/加试/.test(clause);
            if (!publication && !certificate && !language) continue;
            const alternatives = publication ? [clause] : splitOutside(clause.replace(/^或/, ''), /或/);
            const results = alternatives.map(value => publication ? paperCondition(value, papers) : certificateCondition(value, entries));
            if (/^或/.test(clause) && groups.length) groups[groups.length - 1].push(...results);
            else groups.push(results);
        }
        const requirements = groups.map(options => options.find(item => item.status === 'match') || options.find(item => item.status === 'review') || options.find(item => item.status === 'missing') || options[0]);
        const unresolved = requirements.filter(item => item.status !== 'match');
        const mismatch = unresolved.length > 0;
        return { status: mismatch ? 'mismatch' : unresolved.length ? 'review' : 'match', reviews: unresolved.map(item => item.detail), requirements };
    }
    function matchJob(job, profile) {
        const global = globalConditions(job, profile);
        if (global.status === 'mismatch') return { status: 'mismatch', matches: [], reviews: global.reviews, requirements: [] };
        const matches = [];
        profile.educations.forEach((education, index) => {
            const level = educationMatch(job.ed, education.level);
            if (level === 'mismatch') return;
            const major = majorMatch(job.mj, education);
            if (major.status === 'mismatch') return;
            const eligibility = eligibilityConditions(job, profile, education);
            if (eligibility.status === 'mismatch') return;
            const reviews = [...global.reviews];
            if (level === 'review') reviews.push('学历表述需核对');
            if (major.status === 'review') reviews.push(major.reason);
            const other = text(job.ot), description = text(job.pd);
            const degree = degreeConditions(other, education, profile);
            if (degree.status === 'mismatch') return;
            reviews.push(...degree.reviews);
            const residency = residencyRules.match(job, education, profile);
            if (residency.status === 'mismatch') return;
            if (residency.status === 'review') reviews.push(residency.reason);
            const credentials = credentialConditions(other, education, profile);
            if (credentials.status === 'mismatch') return;
            reviews.push(...credentials.reviews);
            const graduateText = other + ' ' + description;
            if (/应届毕业生/.test(graduateText)) {
                if (/不限.{0,8}应届|应届.{0,8}不限/.test(graduateText)) { /* 岗位明确不限应届身份。 */ }
                else if (/应届.{0,15}或|或.{0,15}应届/.test(graduateText)) {
                    if (!['yes', 'regular'].includes(education.graduate)) return;
                    reviews.push('应届条件含并列或特殊表述，请核对原文');
                }
                else {
                    const requiresGraduate = !/非应届毕业生|不含应届毕业生/.test(graduateText);
                    if (['yes', 'regular'].includes(education.graduate) !== requiresGraduate) return;
                }
            }
            matches.push({ index, level: LABELS[education.level], major: education.major, status: reviews.length ? 'review' : 'match', reason: major.reason, reviews: [...new Set(reviews)], requirements: credentials.requirements, residency });
        });
        const confirmed = matches.filter(item => item.status === 'match');
        const selected = confirmed.length ? confirmed : matches;
        return { status: confirmed.length ? 'match' : matches.length ? 'review' : 'mismatch', matches: selected, reviews: confirmed.length ? [] : [...new Set(selected.flatMap(item => item.reviews))], requirements: selected.flatMap(item => item.requirements) };
    }
    return Object.freeze({ LEVELS, LABELS, normalizeProfile, educationMatch, majorMatch, matchJob });
}));
