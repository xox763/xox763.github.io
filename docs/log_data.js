// ==UserScript==
// @name           Hero Wars Data Export (Customized)
// @namespace      http://tampermonkey.net/
// @version        0.0.1
// @description    Custom script for processing Hero Wars game data. Inspired by EnterBrain42's original script.
// @author         Ihg
// @match          https://www.hero-wars.com/*
// @require        https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js
// @icon           https://heroesweb-a.akamaihd.net/i/hw-web/v2/1532007/images/pwa/favicon-16x16.png
// @grant          GM.setValue
// @grant          GM.getValue
// @grant          GM_registerMenuCommand
// @grant          GM_setClipboard
// @grant          unsafeWindow
// @license        MIT
// ==/UserScript==

(function() {
    'use strict';

    // #region Hooking XMLHttpRequest.open

    (function(open) {
        XMLHttpRequest.prototype.open = function() {
            this.addEventListener("readystatechange", function() {
                if (this.readyState === 4 && this.status === 200 && this.responseURL === "https://heroes-wb.nextersglobal.com/api/") {
                    let jsonResponse = this.response;
                    if (typeof jsonResponse === 'string'){
                        jsonResponse = JSON.parse(jsonResponse);
                    }
                    setTimeout(addResponse, 1, this.responseType, jsonResponse);
                }
            }, false);
            open.apply(this, arguments);
        };
    })(XMLHttpRequest.prototype.open);

    // #endregion

    // #region Helper Functions

    function validateSchema(target, schema) {
        if (schema === null) return true;
        if (target === null) return true;
        if (typeof schema !== "object") {
            let sid;
            switch (typeof target) {
                case 'string':    sid = isNaN(target) ? 's' : 'n'; break;
                case 'number':    sid = 'n'; break;
                case 'boolean':   sid = 'b'; break;
                case 'bigint':    sid = 'i'; break;
                case 'undefined': sid = 'u'; break;
                case 'function':  sid = 'f'; break;
                case 'symbol':    sid = 'y'; break;
                case 'object':    sid = target === null ? '_' : null; break;
                default:          sid = null; // Unknown/Unsupported
            }
            return sid === schema;
        } else if (typeof target !== "object") {
            return false;
        }

        const isArrS = Array.isArray(schema);
        const isArrT = Array.isArray(target);
        if (isArrS !== isArrT) return false;
        if (isArrS) {
            const ial = schema.length === 1;
            if (!ial && schema.length !== target.length) return false;
            const childSchema = schema[0];
            return target.every((child, idx) => {
                return validateSchema(child, ial ? childSchema : schema[idx]);
            });
        } else {
            if (Object.hasOwn(schema, '__ial_object__')) {
                const childSchema = schema.__ial_object__;
                // Every key in the target must match the childSchema
                return Object.keys(target).every(key => {
                    if (isNaN(key)) return false;
                    return validateSchema(target[key], childSchema);
                });
            }

            return Object.keys(schema).every(key => {
                if (!Object.hasOwn(target, key)) return false;
                return validateSchema(target[key], schema[key]);
            })
        }
    }

    // GMT+2
    function getGameDayInfo(date, gmtOffset = 2) {
        const gameDate = addHours(date, -gmtOffset);
        const dayInWeek = gameDate.getUTCDay();
        const { weekId, weekStart, weekEnd } = getAbsoluteWeek(gameDate);
        return { dayInWeek, weekId, weekStart: addHours(weekStart, gmtOffset), weekEnd: addHours(weekEnd, gmtOffset) };
    }

    // startDay: sun(0), mon(1), ...
    function getAbsoluteWeek(date, startDay = 1) {
        const MS_PER_DAY = 86400000;
        const MS_PER_WEEK = 604800000;
        const offset = (4 - startDay) * MS_PER_DAY;
        const weekId = Math.floor((date.getTime() + offset) / MS_PER_WEEK);
        const weekStart = new Date(weekId * MS_PER_WEEK - offset);
        const weekEnd = new Date((weekId + 1) * MS_PER_WEEK - offset);
        return { weekId, weekStart, weekEnd };
    }

    function addDays(date, days) {
        const MS_PER_DAY = 86400000; // 24 * 60 * 60 * 1000
        return new Date(date.getTime() + (days * MS_PER_DAY));
    }
    
    function addHours(date, hours) {
        const MS_PER_HOUR = 3600000; // 60 * 60 * 1000;
        return new Date(date.getTime() + (hours * MS_PER_HOUR));
    }

    function addSeconds(date, seconds) {
        const MS_PER_SECOND = 1000;
        return new Date(date.getTime() + (seconds * MS_PER_SECOND));
    }

    async function loadFile(types) {
        try {
            const win = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
            const [fileHandle] = await win.showOpenFilePicker({
                multiple: false,
                excludeAcceptAllOption: false,
                types
            });
            const file = await fileHandle.getFile();
            const content = await file.text();
            return { handle: fileHandle, content };
        } catch (error) {
            if (error.name === 'AbortError') {
                // user cancellation
                return { handle: null, content: null, cancelled: true };
            }
            // Handle errors
            console.error("Load failed:", error);
            return { handle: null, content: null, cancelled: false };
        }
    }

    async function saveFile(suggestedName, content, types) {
        try {
            const win = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
            const showSaveFilePicker = win.showSaveFilePicker.bind(win);
            const handle = await showSaveFilePicker({
                suggestedName,
                types
            });

            const writable = await handle.createWritable();
            await writable.write(content);
            await writable.close();
            return true;
        } catch (error) {
            if (error.name === 'AbortError') {
                // user cancellation
                return true;
            }
            // Handle errors
            console.error("Save failed:", error);
            return false;
        }
    }

    const jsonFileTypes = [{ description: 'JSON Files', accept: { 'text/json': ['.json'] }}];
    const csvFileTypes = [{ description: 'CSV Files', accept: { 'text/csv': ['.csv'] }}];

    async function loadJsonFile() {
        return loadFile(jsonFileTypes);
    }

    async function loadCsvFile() {
        return loadFile(csvFileTypes);
    }

    async function saveJsonFile(suggestedName, content) {
        return saveFile(suggestedName, content, jsonFileTypes);
    }

    async function saveCsvFile(suggestedName, content) {
        return saveFile(suggestedName, content, csvFileTypes);
    }

    function findDataInfo(name, value) {
        if (name && namedDataInfos.has(name))
            return namedDataInfos.get(name);
        if (value !== null) {
            for (const item of dataInfos) {
                if (validateSchema(value, item.schema))
                    return item;
            }
        }
        return null;
    }

    function activityStat(data, membershipState, offset = 0) {
        const stat = data.reverse();
        if (membershipState) {
            const diw = (membershipState.dayInWeek + 6) % 7 + offset; // mon: 0, ..., sun: 6
            if (membershipState.joined) {
                stat.forEach((value, idx) => {
                    if (idx < diw) stat[idx] = null;
                });
            } else {
                stat.forEach((value, idx) => {
                    if (idx > diw) stat[idx] = null;
                });
            }
        }
        return `${stat.join(',')}`;
    }

    function getWinStatus(you, enemy) {
        return you > enemy ? 'Victory' : (you < enemy ? 'Defeat' : 'Draw');
    }

    function getGuildName(name, serverId) {
        return `${name.trim()} Server ${serverId}`;
    }

    function getFinalPrestigeProgress(progress, currentTime, endTime) {
        const _28days = 28 * 24 * 60 * 60 * 1000;
        const prestigeDateInfo = [
            { start: new Date("2024-06-05 02:00:00Z"), end: new Date("2024-06-30 14:00:00Z") },
            { start: new Date("2024-07-01 14:00:00Z"), end: new Date("2024-07-28 14:00:00Z") },
            { start: new Date("2024-07-29 02:00:00Z"), end: new Date("2024-08-25 02:00:00Z") },
            { start: new Date("2024-08-26 02:00:00Z"), end: new Date("2024-09-22 02:00:00Z") },
            { start: new Date("2024-09-23 02:00:00Z"), end: new Date("2024-10-20 02:00:00Z") },
            { start: new Date("2024-10-21 02:00:00Z"), end: new Date("2024-11-17 02:00:00Z") },
            { start: new Date("2024-11-18 02:00:00Z"), end: new Date("2024-12-15 02:00:00Z") },
            { start: new Date("2024-12-16 02:00:00Z"), end: new Date("2025-01-12 02:00:00Z") },
            { start: new Date("2025-01-13 02:00:00Z"), end: new Date("2025-02-09 02:00:00Z") },
            { start: new Date("2025-02-10 02:00:00Z"), end: new Date("2025-03-09 02:00:00Z") },
            { start: new Date("2025-03-10 14:00:00Z"), end: new Date("2025-04-07 14:00:00Z") },
            { start: new Date("2025-04-08 14:00:00Z"), end: new Date("2025-05-06 14:00:00Z") },
            { start: new Date("2025-05-07 14:00:00Z"), end: new Date("2025-06-04 14:00:00Z") },
            { start: new Date("2025-06-05 14:00:00Z"), end: new Date("2025-07-03 14:00:00Z") },
            { start: new Date("2025-07-04 14:00:00Z"), end: new Date("2025-08-01 14:00:00Z") },
            { start: new Date("2025-08-02 14:00:00Z"), end: new Date("2025-08-30 14:00:00Z") },
            { start: new Date("2025-08-31 14:00:00Z"), end: new Date("2025-09-28 14:00:00Z") },
            { start: new Date("2025-09-29 14:00:00Z"), end: new Date("2025-10-27 14:00:00Z") },
            { start: new Date("2025-10-28 14:00:00Z"), end: new Date("2025-11-25 14:00:00Z") },
            { start: new Date("2025-11-26 14:00:00Z"), end: new Date("2025-12-24 14:00:00Z") },
            { start: new Date("2025-12-25 14:00:00Z"), end: new Date("2026-01-26 14:00:00Z") },
        ];

        let startTime = null;
        for (const { start, end } of prestigeDateInfo) {
            if (endTime <= end) {
                startTime = start;
                break;
            }
        }
        if (startTime === null) {
            startTime = new Date(endTime - _28days);
        }

        return Math.floor(progress * _28days / (currentTime - startTime));
    }

    function getPrestigeLevel(progress) {
        const points = [
            0, 5000, 10500, 18000, 27000, 37500, 49750, 63500,
            78750, 95500, 113500, 133000, 154000, 176250, 199750, 224750, 251000,
            278500, 307250, 337500, 368750, 401250, 435250, 470250, 506500, 544000,
            582750, 622500, 663500, 705750, 749250, 793750, 839500, 886500, 934500,
            983750, 1034000, 1085000, 1138000, 1191500, 1246500, 1302250, 1359250, 1417250,
            1476500, 1536750, 1598250, 1660500, 1724000, 1788750, 1854250, 1921000, 1988750,
            2057500, 2127500, 2198500, 2270500, 2343500, 2417500, 2492500, 2568750, 2645750,
            2724000, 2803250, 2883500, 2964750, 3047000, 3130250, 3214500, 3300000
        ];
        let level = null;
        for (const [idx, cap] of points.entries()) {
            if (progress < cap) {
                level = idx;
                break;
            }
        }
        if (level === null) {
            level = Math.floor((progress - 3300000) / 90000) + 70;
        }
        return level;
    }

    // #endregion

    // #region Data & States

    const dataInfos = [
        {
            id: "clanGetInfo",
            description: "Guild Info",
            collectionPath: "Home",
            named: true,
            needed: true,
            schema: {
                "clan": null,
                "membersStat": null,
                "stat": null,
                "serverResetTime": "n",
                "clanWarEndSeasonTime": "n",
                "freeClanChangeInterval": null,
                "giftUids": null
            },
        },
        {
            id: "clanGetLog",
            description: "Guild Status",
            collectionPath: "Home > Daily Quests > Guild Quests > Status",
            named: true,
            needed: true,
            schema: {
                "history": null,
                "users": null
            },
        },
        {
            id: "guildStats",
            description: "Guild Stats",
            collectionPath: "Home > Daily Quests > Guild Quests > Statistics",
            named: false,
            needed: true,
            schema: {
                "stat": null,
                "today": "n",
                "dayInWeek": "n",
                "giftsCount": "n"
            }
        },
        {
            id: "clan_prestigeGetInfo",
            description: "Guild Prestige Progress",
            collectionPath: "Home > Daily Quests > Guild Quests > Prestige",
            named: true,
            needed: true,
            schema: {
                "prestigeId": "n",
                "prestigeCount": "n",
                "userPrestigeCount": "n",
                "farmedPrestigeLevels": null,
                "endTime": "n",
                "prestigeStartPopupViewed": "b",
                "nextTime": "n"
            }
        },
        {
            id: "clanWarGetInfo",
            description: "Guild War Info",
            collectionPath: "Home",
            named: true,
            needed: true,
            schema: {
                "season": "n",
                "day": "n",
                "endTime": "n",
                "nextWarTime": "n",
                "nextLockTime": "n",
                // "avgLevel": "n",
                // "league": "n",
                // "enemyId": "n",
                // "enemyClan": null,
                // "enemyClanMembers": null,
                // "points": "n",
                // "enemyPoints": "n",
                // "clanTries": null,
                // "enemyClanTries": null,
                // "myTries": "n",
                // "enemySlots": null,
                // "ourSlots": null,
                // "arePointsMax": "b",
            },
        },
        {
            id: "clanWarsLog",
            description: "Guild Wars Log",
            collectionPath: "Home > Guild > Guild War > Guild War > Log",
            named: false,
            needed: true,
            schema: {
                "history": null,
                "results": null,
            }
        },
        {
            id: "clanWarLeaderboard",
            description: "Guild War Leaderboard",
            collectionPath: "Home > Guild > Guild War > Guild War > Leagues > Previous week",
            named: false,
            needed: true,
            schema: {
                "top": null,
                "promoCount": "n",
                "clans": null,
            }
        },
        {
            id: "crossClanWar_getInfo",
            description: "Clash of Worlds Info",
            collectionPath: "Home > Guild > Guild War > Clash of Worlds",
            named: true,
            needed: false,
            schema: {
                "nextWarTime": "n",
                "nextLockTime": "n",
                "plannedSeason": "n",
                "season": "n",
                "seasonEndTime": "n",
                "nextSeasonStartTime": "n",
                "requiredDefendedSlots": "n",
                "defendedSlots": "n",
                "settings": null,
                "war": null,
                "rating": "n",
                "division": "n",
                "league": "n",
                "maxLeague": "n"
            }
        },
        {
            id: "crossClanWarLog",
            description: "Clash of Worlds Log",
            collectionPath: "Home > Guild > Guild War > Clash of Worlds > Log",
            named: false,
            needed: true,
            schema: [{
                "season": "n",
                "war": "n",
                "ctime": "n",
                "enemyClan": {
                    "id": "n",
                    "serverId": "n",
                    "title": "s",
                    "icon": null
                },
                "ratingDelta": "n",
                "points": "n",
                "enemyPoints": "n",
                "rating": "n"
            }]
        },
        {
            id: "crossClanWarBattleLog",
            description: "Clash of Worlds Battle Log",
            collectionPath: "Home > Guild > Guild War > Clash of Worlds > Log > More",
            named: false,
            needed: true,
            accumulation: true,
            schema: {
                "attack": null,
                "defence": null,
                "users": null
            }
        },
        {
            id: "clanRaid_getInfo",
            description: "Guild Raid Info",
            collectionPath: "Home > Guild > Asgard > Guild Raid",
            named: true,
            needed: true,
            schema: {
                "boss": null,
                "nodes": null,
                "shop": null,
                "buffs": null,
                "flags": null,
                "stats": null,
                "userStats": null,
                "attempts": "n",
                "bossAttempts": "n",
                "lastBossId": "n",
                "coins": "n"
            }
        },
        {
            id: "clanRaidMemberInfo",
            description: "Guild Raid Member Info",
            collectionPath: "Home > Guild > Asgard > Guild Raid > Log",
            named: false,
            needed: true,
            schema: {
                __ial_object__: {
                    "hasActiveSubscription": "b",
                    "bonusClanBuffPoints": "n", // morale/2
                    "raidAvailable": "b"
                }
            }
        },
        {
            id: "clanRaidBriefStats",
            description: "Guild Raid Damage Dealt to Boss & Morale Points",
            collectionPath: "Home > Guild > Asgard > Guild Raid > Log",
            named: false,
            needed: true,
            schema: {
                __ial_object__: {
                    "bossDamage": "n",
                    "nodesPoints": null, // morale
                    "nodesAttemptsSpent": "n",
                    "bossAttemptsSpent": "n"
                }
            }
        },
        {
            id: "clanRaidBossLog",
            description: "Guild Raid Boss Log",
            collectionPath: "Home > Guild > Asgard > Guild Raid > Log",
            named: false,
            needed: true,
            schema: {
                __ial_object__: {
                    __ial_object__: {
                        "result": {
                            "damage": null
                        },
                        "userId": "n",
                        "typeId": "n",
                        "attackers": null,
                        "defenders": null,
                        "effects": null,
                        "reward": null,
                        "startTime": "n",
                        "seed": "n",
                        "type": "s",
                        "id": "n",
                        "progress": null,
                        "endTime": "n"
                    }
                }
            }
        },
        {
            id: "clanRaidMinionLog",
            description: "Guild Raid Minion Log",
            collectionPath: "Home > Guild > Asgard > Guild Raid > Log",
            named: false,
            needed: true,
            schema: {
                __ial_object__: {
                    __ial_object__: {
                        "result": {
                            "points": "n"
                        },
                        "userId": "n",
                        "typeId": "n",
                        "attackers": null,
                        "defenders": null,
                        "effects": null,
                        "reward": null,
                        "startTime": "n",
                        "seed": "n",
                        "type": "s",
                        "id": "n",
                        "progress": null,
                        "endTime": "n"
                    }
                }
            }
        },
    ];

    const namedDataInfos = new Map(dataInfos.filter(data => data.named).map(data => [data.id, data]));

    const fortInfo = [
        { name: "Mage Academy",         type: "Hero",  startSlotId: 1,    numSlots: 3 },
        { name: "Lighthouse",           type: "Hero",  startSlotId: 5,    numSlots: 5 },
        { name: "Barracks",             type: "Hero",  startSlotId: 12,   numSlots: 3 },
        { name: "Bridge",               type: "Titan", startSlotId: 15,   numSlots: 6 },
        { name: "Engineerium",          type: "Hero",  startSlotId: 21,   numSlots: 5 },
        { name: "Spring of Elements",   type: "Titan", startSlotId: 26,   numSlots: 4 },
        { name: "Foundry",              type: "Hero",  startSlotId: 30,   numSlots: 5 },
        { name: "Gates of Nature",      type: "Titan", startSlotId: 35,   numSlots: 4 },
        { name: "Bastion of Fire",      type: "Titan", startSlotId: 39,   numSlots: 4 },
        { name: "Bastion of Ice",       type: "Titan", startSlotId: 43,   numSlots: 4 },
        { name: "Ether Prism",          type: "Titan", startSlotId: 47,   numSlots: 5 },
        { name: "Shooting Range",       type: "Hero",  startSlotId: 52,   numSlots: 5 },
        { name: "Bastion",              type: "Hero",  startSlotId: 57,   numSlots: 5 },
        { name: "Altar of Life",        type: "Titan", startSlotId: 62,   numSlots: 5 },
        { name: "Heroes' Bridge",       type: "Hero",  startSlotId: 67,   numSlots: 6 },
        { name: "Alchemy Tower",        type: "Hero",  startSlotId: 73,   numSlots: 5 },
        { name: "City Hall",            type: "Hero",  startSlotId: 78,   numSlots: 5 },
        { name: "Sun Temple",           type: "Titan", startSlotId: 83,   numSlots: 4 },
        { name: "Moon Temple",          type: "Titan", startSlotId: 87,   numSlots: 4 },
        { name: "Citadel",              type: "Hero",  startSlotId: 91,   numSlots: 8 }
    ];

    const fortsById = fortInfo.reduce((acc, fort) => {
        const { startSlotId, numSlots } = fort;
        for (let i = 0; i < numSlots; i++) {
            acc[startSlotId + i] = fort;
        }
        return acc;
    }, []);

    const rawData = new Map();
    const loggedData = [];
    let dataLoggingEnabled = false;

    // #endregion

    // #region Response Handlers

    function addResponse(responseType, response) {
        if (responseType !== ""
            || !Object.hasOwn(response, 'results')
            || !Array.isArray(response.results)) {
            console.warn('informal data', {responseType, response});
            return;
        }

        const timestamp = new Date();
        const { dayInWeek, weekId } = getGameDayInfo(timestamp); // [dayInWeek, weekId] sun: 0, ...

        response.results.forEach(item => {
            if (!Object.hasOwn(item, "ident")
                || !Object.hasOwn(item, "result")
                || !Object.hasOwn(item.result, "response")
                // || typeof item.result.response !== 'object'
                // || item.result.response === null
            ) {
                console.warn('informal data', item);
                return;
            }

            const itemName = item.ident;
            const itemValue = item.result.response;
            let dataInfo;
            if (itemName.startsWith("group_") || itemName === "body") {
                dataInfo = findDataInfo(null, itemValue);
            } else {
                dataInfo = findDataInfo(itemName, null);
            }

            if (dataInfo) {
                const data = {
                    name: dataInfo.id,
                    description: dataInfo.description,
                    datetime: { timestamp, dayInWeek, weekId },
                    data: itemValue
                };
                if (dataInfo.accumulation) {
                    if (rawData.has(dataInfo.id)) {
                        const cont = rawData.get(dataInfo.id);
                        cont.data.push(itemValue);
                    } else {
                        data.data = [itemValue];
                        rawData.set(dataInfo.id, data);
                    }
                } else {
                    rawData.set(dataInfo.id, data);
                }
                console.log(`👉 ${dataInfo.description}`);
            } else {
                if (dataLoggingEnabled)
                    console.log(`👉 <${itemName}>`);
            }

            // data logging
            if (dataLoggingEnabled)
                loggedData.push([itemName, itemValue]);
        });
    }

    async function saveStatsCSV(collectedData, collectionTime) {
        const data = Object.fromEntries(collectedData);
        const { dayInWeek, weekId, weekStart, weekEnd } = getGameDayInfo(collectionTime);
        const weekDateStrings = [0, 1, 2, 3, 4, 5, 6].map(days => addDays(weekStart, days).toISOString().slice(0,10));
        const dateName = weekDateStrings[0];
        const members = new Map();
        const loggedMembers = new Map();

        // #region Validations
        // check for not collected data
        const notCollected = [];
        for (const {id, needed, description} of dataInfos) {
            if (needed) {
                if (!collectedData.has(id)) {
                    notCollected.push(description);
                } else {
                    const itemData = collectedData.get(id).data;
                    if (Array.isArray(itemData) && itemData.length === 0) {
                        notCollected.push(description);
                    }
                }
            }
        }
        if (notCollected.length) {
            alert(`The following data has not been collected:\n\n- ${notCollected.join('\n- ')}`);
            return;
        }

        // check for collected day and time
        const invalidEntries = [];
        for (const [, { description, datetime }] of collectedData) {
            const isSundayCollection = datetime.dayInWeek === 0;
            const isCurrentWeekSunday = (datetime.weekId === weekId && dayInWeek === 0);
            const isNextWeekMonday = (datetime.weekId + 1 === weekId && dayInWeek === 1);
            const isValidDatetime = isSundayCollection && (isCurrentWeekSunday || isNextWeekMonday);
            if (!isValidDatetime)
                invalidEntries.push(description);
        }
        if (invalidEntries.length) {
            const errorMessage = `The following data was not collected on Sunday or belongs to a different week:\n\n- ${invalidEntries.join('\n- ')}\n\nDo you still want to proceed?`;
            if (!confirm(errorMessage)) return;
        }

        // check for guild war leaderboard
        const gw = data.clanWarsLog.data;
        const gwResult = gw.results.previous;
        const gwLeaderboard = data.clanWarLeaderboard.data;
        const gwl = gwLeaderboard.top[0];
        const gwLeague = Number(gwResult.league);
        if (Number(gwl.points) === 0 || Number(gwl.league) !== gwLeague) {
            alert("You must refer to the league page of the guild from the previous week on the Guild War Leaderboard.");
            return;
        }

        // #endregion

        // #region Guild Stats
        // create members info
        const guildInfo = data.clanGetInfo.data.clan;
        for (const [id, member] of Object.entries(guildInfo.members)) {
            member.warrior = false;
            members.set(Number(id), member);
        }

        // create logged members info
        const guildLog = data.clanGetLog.data;
        for (const [id, member] of Object.entries(guildLog.users)) {
            loggedMembers.set(Number(id), member);
        }

        // add join or kick datetime
        const weekStartCtime = weekStart.getTime() / 1000;
        const weekEndCtime = weekEnd.getTime() / 1000;
        const eventHistory = guildLog.history.filter(({ ctime }) =>
            ctime >= weekStartCtime && ctime < weekEndCtime            
        );
        eventHistory.forEach(({ userId, event, ctime, details }) => {
            if (event === "join" || event === "leave" || event === "autokick" || event === "kick") {
                const member = members.get(Number(event !== "kick" ? userId : details.userId));
                if (member != null) {
                    const date = new Date(ctime * 1000);
                    const { dayInWeek } = getGameDayInfo(date);
                    member.membershipState = { event, joined: event === "join", dayInWeek, date, kickBy: event === "kick" ? userId : null };
                }
            }
        });

        // add guild war champion flag
        guildInfo.warriors.forEach(id => {
            const member = members.get(Number(id));
            if (member != null)
                member.warrior = true;
        });

        // guild name
        const guildName = getGuildName(guildInfo.title, guildInfo.serverId);

        // add activity stats
        data.guildStats.data.stat.forEach(({id, activity, dungeonActivity: titanite, adventureStat: adventures, clanWarStat: guildWar, prestigeStat: guildPrestige, clanGifts: gifts}) => {
            const member = members.get(Number(id));
            if (member != null) {
                member.stats = { activity, titanite, adventures, guildWar, guildPrestige, gifts };
            }
        });

        // add clash of worlds stats
        const cowLog = data.crossClanWarLog.data
            .filter(({ctime}) => {
                ctime = Number(ctime);
                return ctime >= weekStartCtime && ctime < weekEndCtime;
            })
            .sort((a, b) => b.season * 1000 + b.war - (a.season * 1000 + a.war))
            .slice(0, 2)
            .reverse();
        const cowBattleLog = new Map();
        data.crossClanWarBattleLog.data.forEach(data => {
            Object.values(data.users).every(({ serverId, clanTitle: title }) => {
                const enemyGuildName = getGuildName(title, serverId);
                if (enemyGuildName !== guildName) {
                    cowBattleLog.set(enemyGuildName, data);
                    return false;
                }
                return true;
            })
        });
        cowLog.forEach(({ war, enemyClan: { serverId, title } }) => {
            const enemyGuildName = getGuildName(title, serverId);
            if (!cowBattleLog.has(enemyGuildName)) {
                alert(`The following data has not been collected:\n\n- Clash of Worlds Battle Log (for ${enemyGuildName})`);
                return;
            }
            const battleDay = (war % 2 !== 0 ? 3 : 0); // THU, SUN
            const battleLog = new Map();
            cowBattleLog.get(enemyGuildName).attack
                .forEach(({ attackerId, slotId }) => {
                    if (battleLog.has(attackerId)) {
                        battleLog.get(attackerId).push(slotId);
                    } else {
                        battleLog.set(attackerId, [slotId]);
                    }
                });

            for (const [id, member] of members) {
                const attackSlots = battleLog.get(id);
                const counts = attackSlots?.reduce((acc, slotId) => {
                    return acc + (fortsById[slotId].type === "Hero" ? 1 : 100);
                }, 0) ?? 0;
                const heroCount = counts % 100;
                const titanCount = Math.floor(counts / 100);
                if (!member.stats.cowHero) {
                    member.stats.cowHero = [,,,,,,,];
                    member.stats.cowTitan = [,,,,,,,];
                }
                member.stats.cowHero[battleDay] = heroCount;
                member.stats.cowTitan[battleDay] = titanCount;
            }
        });

        // add raid stats
        // hasActiveSubscription, bonusClanBuffPoints, raidAvailable
        for (const [id, value] of Object.entries(data.clanRaidMemberInfo.data)) {
            const member = members.get(Number(id));
            if (member != null) {
                member.raid = value;
            }
        }
        // set the missing attr, raidAvailable, for collector
        for (const [id, member] of members) {
            if (!Object.hasOwn(member, 'raid'))
                member.raid = { hasActiveSubscription: null, bonusClanBuffPoints: null, raidAvailable: true };
        }

        // bossDamage, nodesPoints, nodesAttemptsSpent, bossAttemptsSpent
        for (const [id, value] of Object.entries(data.clanRaidBriefStats.data)) {
            const member = members.get(Number(id));
            if (member != null && member.raid.raidAvailable) {
                member.raid = { ...member.raid, ...value };
            }
        }
        // set the missing attrs for raidAvailable is true
        for (const [id, member] of members) {
            if (!member.raid.raidAvailable) continue;
            if (!Object.hasOwn(member.raid, 'bossDamage'))
                member.raid.bossDamage = 0;
            if (!Object.hasOwn(member.raid, 'nodesPoints'))
                member.raid.nodesPoints = 0;
            if (!Object.hasOwn(member.raid, 'nodesAttemptsSpent'))
                member.raid.nodesAttemptsSpent = 0;
            if (!Object.hasOwn(member.raid, 'bossAttemptsSpent'))
                member.raid.bossAttemptsSpent = 0;
        }
        // type nomalization
        members.forEach(member => {
            if (member.raid) {
                Object.entries(member.raid).forEach(([key, value]) => {
                    if (typeof value === "string" && !isNaN(value))
                        member.raid[key] = Number(value);
                });
            }
        });

        // add boss log
        // bossLog
        for (const [id, value] of Object.entries(data.clanRaidBossLog.data)) {
            const member = members.get(Number(id));
            if (member != null) {
                const bossLog = [];
                for (const { attackers, effects, result } of Object.values(value)) {
                    bossLog.push({ attackers, effects, result });
                }
                member.raid.bossLog = bossLog;
            }
        }

        // output stats.csv
        const statsCsvData = [];
        for (const [id, member] of members) {
            const joinDate = member.membershipState?.joined
                            ? addHours(member.membershipState.date, -2).toISOString().slice(0, 10)
                            : null;

            statsCsvData.push([
                id,
                member.name,
                member.level,
                joinDate,
                member.lastLoginTime,
                member.warrior,
                activityStat(member.stats.titanite, member.membershipState),
                activityStat(member.stats.activity, member.membershipState),
                activityStat(member.stats.guildPrestige, member.membershipState),
                activityStat(member.stats.adventures, member.membershipState),
                member.warrior ? activityStat(member.stats.guildWar, member.membershipState) : ',,,,,,',
                activityStat(member.stats.cowHero, member.membershipState, 4),
                activityStat(member.stats.cowTitan, member.membershipState, 4),
                activityStat(member.stats.gifts, member.membershipState),
                member.raid.raidAvailable,
                member.raid.nodesAttemptsSpent,
                member.raid.nodesPoints,
                member.raid.bossAttemptsSpent,
                member.raid.bossDamage
            ]);
        }

        const statsCsvString = Papa.unparse({
            fields: ["id", "name", "level", "joinDate", "lastLoginTime", "warrior", "titanite"
                , "activity", "guildPrestige", "adventures", "guildWar", "cowHero", "cowTitan"
                , "gifts", "raidAvailable", "raidMinionAttacks", "raidMorale", "raidBossAttacks"
                , "raidBossDamage"],
            data: statsCsvData
        });
        const statsFileName = `stats_${dateName}.csv`;

        if (!await saveCsvFile(statsFileName, statsCsvString)) {
            GM_setClipboard(statsCsvString);
            const continueProcess = confirm(
                `File saving failed. The content of "${statsFileName}" has been copied to your clipboard instead.\n\nDo you want to proceed with the next file?`
            );

            if (!continueProcess) {
                return; // Stop execution if the user chooses 'Cancel'
            }
        }

        // #endregion

        // #region Event Log
        // add guild war log
        const eventLogCsvData = gw.history.slice(-5).map(({ day, points, enemyClan: { title: enemyGuildName }, enemyPoints }) => {
            day = Number(day) - 1;
            points = Number(points);
            enemyPoints = Number(enemyPoints);
            return [
                weekDateStrings[day],                           // Date
                null,                                           // Days
                'Guild War',                                    // Type
                enemyGuildName.trim(),                          // Name
                getWinStatus(points, enemyPoints),              // Status
                points,                                         // Value1
                enemyPoints,                                    // Value2
                null,                                           // Value3
                null                                            // Value4
            ];
        });

        // add guild war result
        const gwPosition = Number(gwResult.position);
        const gwPoints = Number(gwLeaderboard.top[gwPosition-1].points);
        const gapToAbove = Number(gwLeaderboard.top[gwPosition-2]?.points ?? gwPoints) - gwPoints;
        const gapToBelow = Number(gwLeaderboard.top[gwPosition]?.points ?? gwPoints) - gwPoints;
        const gwInfo = data.clanWarGetInfo.data;
        const gwWeek = `Week ${Number(gwInfo.season.slice(-2))}`;
        const gwLeagues = ['Gold League', 'Silver League', 'Bronze League', 'Qualifying League'];
        eventLogCsvData.push([
            weekDateStrings[5],         // Date
            null,                       // Days
            'Guild War - Result',       // Type
            gwWeek,                     // Name
            gwLeagues[gwLeague - 1],    // Status
            gwPosition,                 // Value1
            gwPoints,                   // Value2
            gapToAbove,                 // Value3
            gapToBelow                  // Value4
        ]);

        // add clash of worlds log
        let cowWeek = null;
        cowLog.forEach(({ war, enemyClan: { serverId, title }, ratingDelta, points, enemyPoints }, idx) => {
            points = Number(points);
            enemyPoints = Number(enemyPoints);
            cowWeek = cowWeek ?? Math.floor((Number(war) + 1) / 2);
            eventLogCsvData.push([
                weekDateStrings[idx * 3],                       // Date
                null,                                           // Days
                'Clash of Worlds',                              // Type
                getGuildName(title, serverId),                  // Name
                getWinStatus(points, enemyPoints),              // Status
                points,                                         // Value1
                enemyPoints,                                    // Value2
                ratingDelta,                                    // Value3
                null                                            // Value4
            ]);
        });

        // add clash of worlds result
        if (cowLog.length) {
            const cowInfo = data.crossClanWar_getInfo.data;
            // const cowLeagues = ['Baron League', 'Viscount League', 'Earl League', 'Marquis League', 'Duke League'];
            const cowLeagues = ['Duke League', 'Marquis League', 'Earl League', 'Viscount League', 'Baron League'];
            const league = Number(cowInfo.league) - 1;
            eventLogCsvData.push([
                weekDateStrings[6],                             // Date
                null,                                           // Days
                'Clash of Worlds - Result',                     // Type
                `Season ${cowInfo.season}. Week ${cowWeek}`,    // Name
                cowLeagues[league],                             // Status
                cowInfo.division - league * 5,                  // Value1
                cowInfo.rating,                                 // Value2
                null,                                           // Value3
                null                                            // Value4
            ]);
        }

        // add raid result
        const raidInfo = data.clanRaid_getInfo.data;
        const bossId = Number(raidInfo.stats.currentBoss) - 1;
        const guildMorale = raidInfo.stats.points;
        const bossKilled = Object.entries(raidInfo.stats.bossKilled);
        const memberArr = Array.from(members.values());
        const bossDamage = memberArr.reduce((acc, member) => acc + (member.raid?.bossDamage ?? 0), 0);
        const bossAttemptsSpent = memberArr.reduce((acc, member) => acc + (member.raid?.bossAttemptsSpent ?? 0), 0);
        const bossMaxLevelKilled = bossKilled.reduce((acc, [level, count]) => {
            if (count) {
                level = Number(level);
                acc = Math.max(acc, Number(level));
            }
            return acc;
        }, 0);
        const raidTypes = ['Cradle of the Stars', 'The Phantom Orchestra'];
        eventLogCsvData.push([
            weekDateStrings[6],                         // Date
            null,                                       // Days
            'Raid - Result',                            // Type
            raidTypes[bossId],                          // Name
            bossMaxLevelKilled ?  'Victory' : 'Defeat', // Status
            bossMaxLevelKilled,                         // Value1
            guildMorale,                                // Value2
            bossDamage,                                 // Value3
            bossAttemptsSpent                           // Value4
        ]);

        // add raid difficulty
        const raidSetLevel = eventHistory.filter(({ event }) => event === "clanRaidSetLevel");
        const nextLevel = raidSetLevel.length ? raidSetLevel[0].details.level : null;
        eventLogCsvData.push([
            weekDateStrings[6],                         // Date
            null,                                       // Days
            'Raid Difficulty',                          // Type
            raidTypes[1 - bossId],                      // Name
            null,                                       // Status
            nextLevel,                                  // Value1
            null,                                       // Value2
            null,                                       // Value3
            null                                        // Value4
        ]);

        // add guild info
        const giftsCount = guildInfo.giftsCount;
        const prestigeProgress = data.clan_prestigeGetInfo.data.prestigeCount;
        const finalPrestigeProgress = getFinalPrestigeProgress(
            prestigeProgress
            , collectionTime
            , new Date(data.clan_prestigeGetInfo.data.endTime * 1000)
        );
        const finalPrestigeLevel = getPrestigeLevel(finalPrestigeProgress);
        eventLogCsvData.push([
            weekDateStrings[6],     // Date
            null,                   // Days
            'Guild Info',           // Type
            guildName,              // Name
            null,                   // Status
            prestigeProgress,       // Value1 - prestige progress
            finalPrestigeLevel,     // Value2 - final prestige level
            giftsCount,             // Value3 - gifts count
            null                    // Value4
        ]);

        // add membership info
        eventHistory.reverse().forEach(({ userId, event, ctime, details }) => {
            let status = null;
            let message = null;
            let targetUserId = userId;
            let performerId = null;
            switch(event) {
                case "join":
                    status = "Joined";
                    break;
                case "leave":
                    status = "Left";
                    break;
                case "autokick":
                    status = "Kicked";
                    performerId = "<system>";
                    message = "Automatically dismissed due to inactive player settings.";
                    break;
                case "kick":
                    const blackList = eventHistory.find(({event}) => event.startsWith("blackList"));
                    if (blackList && blackList.event === "blackListAdd") {
                        status = "Banned";
                        message = "Expelled due to inactivity and blacklisted.";
                    } else {
                        status = "Kicked";
                        message = "Expelled due to long-term inactivity.";
                    }
                    targetUserId = details.userId;
                    performerId = userId;
                    break;
                default:
                    return;
            }
            const { dayInWeek } = getGameDayInfo(new Date(ctime * 1000));
            const dateStr = weekDateStrings[(dayInWeek + 6) % 7];
            const member = members.get(Number(targetUserId)) ?? loggedMembers.get(Number(targetUserId));
            const name = member?.name ?? userId;

            eventLogCsvData.push([
                dateStr,                // Date
                null,                   // Days
                'Membership Update',    // Type
                name,                   // Name
                status,                 // Status
                targetUserId,           // Value1
                performerId,            // Value2
                message,                // Value3
                null,                   // Value4
            ]);
        });

        const membershipList = new Map();
        for (const [id, member] of loggedMembers) {
            membershipList.set(id, member.name);
        }
        for (const [id, member] of members) {
            membershipList.set(id, member.name);
        }
        console.log(membershipList);
        Array.from(membershipList)
            .sort((a, b) => b[1].localeCompare(a[1]))
            .forEach(([id, name]) => {
                eventLogCsvData.push([
                    weekDateStrings[6],     // Date
                    null,                   // Days
                    'Membership Update',    // Type
                    name,                   // Name
                    'ID',                   // Status
                    id,                     // Value1
                    null,                   // Value2
                    null,                   // Value3
                    null,                   // Value4
                ]);
            });

        const eventLogCsvString = Papa.unparse({
            fields: ["Date", "Days", "Type", "Name", "Status", "Value1", "Value2", "Value3", "Value4"],
            data: eventLogCsvData
        });
        const eventLogFileName = `event_log_${dateName}.csv`;

        if (!await saveCsvFile(eventLogFileName, eventLogCsvString)) {
            alert(`File saving failed: ${eventLogFileName}`);
            return;
        }

        // #endregion

        // #region Raid Boss Log

        // member.raid
        // - { hasActiveSubscription, bonusClanBuffPoints, raidAvailable, bossDamage, nodesPoints, nodesAttemptsSpent, bossAttemptsSpent }
        // member.raid.bossLog
        // - [{ attackers, effects, result }]

        const bossLogCsvData = [];
        for (const [id, member] of members) {
            if (member.raid?.bossLog) {
                // add id, name
                member.raid.bossLog.forEach(({ attackers, effects, result }) => {
                    const record = [
                        id,
                        member.name,
                    ];
    
                    let teamPower = 0;
                    // add slots: [id, power, petId]
                    const teamInfo = Object.values(attackers).map(({ id, level, color: promotion, star, power, petId }) => {
                        teamPower += power;
                        return `${id},${level},${promotion},${star},${power},${petId ? petId : ''}`;
                    });
                    record.push(...teamInfo);

                    // add slot padding
                    if (teamInfo.length < 6) {
                        record.push(...Array(6 - teamInfo.length).fill(null));
                    }

                    // add team power, bossLevel, damage
                    const damage = Object.values(result.damage).reduce((acc, value) => acc + value, 0);
                    record.push(teamPower, result.level, damage);

                    // add buffs
                    record.push(JSON.stringify(effects.attackers));

                    // add record
                    bossLogCsvData.push(record);
                });
                
            }
        }

        const bossLogCsvString = Papa.unparse({
            fields: ["id", "name", "slot1", "slot2", "slot3", "slot4", "slot5", "slot6"
                , "power", "bossLevel", "damage", "buffs"],
            data: bossLogCsvData
        });
        const bossLogFileName = `boss_log_${dateName}.csv`;

        if (!await saveCsvFile(bossLogFileName, bossLogCsvString)) {
            alert(`File saving failed: "${bossLogFileName}"`);
            return;
        }

        // #endregion
    }

    function updateSundayStatsCSV(collectedData, statsCsvString) {
        const { data: statsCsv } = Papa.parse(statsCsvString, {
            header: true,
            dynamicTyping: false
        });

        const data = Object.fromEntries(collectedData);
        const members = new Map(statsCsv.map(item => [Number(item.id), item]));

        // #region Validations
        // check for not collected data
        const notCollected = [];
        if (!collectedData.has("guildStats")) {
            notCollected.push("Guild Stats");
        }
        if (notCollected.length) {
            alert(`The following data has not been collected:\n\n- ${notCollected.join('\n- ')}`);
            return null;
        }

        // check for collected day and time
        const invalidEntries = [];
        const datetime = collectedData.get("guildStats").datetime;
        if (datetime.dayInWeek === 0)
            invalidEntries.push(description);
        if (invalidEntries.length) {
            alert(`The following data was collected on Sunday:\n\n- ${invalidEntries.join('\n- ')}`);
            return null;
        }

        // #endregion

        function replaceSundayStat(dataString, value) {
            const data = dataString.split(',');
            data[6] = value;
            return data.join(',');
        }

        const sundayIdx = datetime.dayInWeek; // 6-(6-diw) = diw
        // add activity stats
        data.guildStats.data.stat.map(({id, activity, dungeonActivity: titanite, adventureStat: adventures
                                        , clanWarStat: guildWar, prestigeStat: guildPrestige, clanGifts: gifts}) => {
            const member = members.get(Number(id));
            if (member != null) {
                member.titanite = replaceSundayStat(member.titanite, titanite[sundayIdx]);
                member.activity = replaceSundayStat(member.activity, activity[sundayIdx]);
                member.guildPrestige = replaceSundayStat(member.guildPrestige, guildPrestige[sundayIdx]);
                member.adventures = replaceSundayStat(member.adventures, adventures[sundayIdx]);
                // member.guildWar = replaceSundayStat(member.guildWar, guildWar[sundayIdx]);
                member.gifts = replaceSundayStat(member.gifts, gifts[sundayIdx]);
            }
        });

        // return updated stats.csv
        return Papa.unparse(Array.from(members.values()));
    }

    async function saveDataJSON(collectedData, collectionTime) {
        const { weekStart } = getGameDayInfo(collectionTime);
        const weekDateStrings = [0, 1, 2, 3, 4, 5, 6].map(days => addDays(weekStart, days).toISOString().slice(0,10));

        const dateName = weekDateStrings[0];
        const content = JSON.stringify([...collectedData.values()]);
        await saveJsonFile(`data_${dateName}.json`, content);
    }

    // #endregion

    // #region Menu Commands

    GM_registerMenuCommand('Export Data (CSV)', async () => {
        await saveStatsCSV(rawData, new Date());
    });

    GM_registerMenuCommand('Update Sunday Stats', async () => {
        const { handle, content } = await loadCsvFile();
        if (content != null) {
            const updatedContent = updateSundayStatsCSV(rawData, content);
            if (updatedContent != null) {
                const writable = await handle.createWritable();
                await writable.write(updatedContent);
                await writable.close();
            }
        }
    });

    if (true || debugMenuCommandsEnabled) {
        GM_registerMenuCommand('(Debug) Copy Data.JSON', async () => {
            GM_setClipboard(JSON.stringify([...rawData.values()]));
        });
        GM_registerMenuCommand('(Debug) Save Data.JSON', async () => {
            await saveDataJSON(rawData, new Date());
        });
        GM_registerMenuCommand('(Debug) Load Data.JSON & Save Stats.CSV', async () => {
            const { content } = await loadJsonFile();
            if (content != null) {
                try {
                    const data = new Map();
                    JSON.parse(content).forEach(itemData => {
                        itemData.datetime.timestamp = new Date(itemData.datetime.timestamp);
                        data.set(itemData.name, itemData);
                    });
    
                    const collectionTime = data.get("guildStats").datetime.timestamp;
                    await saveStatsCSV(data, collectionTime);
                } catch (error) {
                    let message;
                    if (error instanceof SyntaxError) {
                        message = `Invalid JSON format: ${error.message}`;
                    } else {
                        message = `An unexpected error occurred: ${error.message}`;
                    }
                    alert(message);
                    console.error(message, error);
                }
            }
        });
        GM_registerMenuCommand('(Debug) Toggle Data Logging', async () => {
            dataLoggingEnabled = !dataLoggingEnabled;
            console.log('Data Logging: ' + (dataLoggingEnabled ? 'Enabled' : 'Disabled'));
        });
        GM_registerMenuCommand('(Debug) Copy Logged Data.JSON', async () => {
            GM_setClipboard(JSON.stringify(loggedData));
        });
        GM_registerMenuCommand('(Debug) Clear Logged Data', async () => {
            loggedData.length = 0;
        });
    }

    // #endregion
})();