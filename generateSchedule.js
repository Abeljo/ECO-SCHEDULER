/**
 *  Auther ABEL JO 
 * 
 * OUR PEST CONTROL SCHEDULER
 * 
 * This system automates the scheduling of pest control visits for March.
 * It strictly adheres to spacing rules, team capacities, and working day constraints.
 * 
 * Frequency Rules:
 * - Monthly: 2 visits (Follow-up 12-16 days after first)
 * - Quarterly: 1 visit in March
 * - Bi-Monthly: 4 visits (7-8 days apart)
 * - VIP_Every_2_Days: Consistent pattern, allows Sunday
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const {
    format,
    addDays,
    isSunday,
    startOfMonth,
    endOfMonth,
    eachDayOfInterval,
    differenceInDays,
    getDay,
    isSameDay
} = require('date-fns');

// Constants
const MARCH_YEAR = 2026; // Current year or specific
const MONTH_INDEX = 2; // March is 0-indexed in JS Date (0=Jan, 1=Feb, 2=Mar)
const CAPACITY_LIMIT = 18; // Max target per day
const EMERGENCY_RESERVE = 1; // 1 hour/job reserved
const NUMBER_OF_CARS = 2;
const TEAMS_PER_CAR = 2;
const MAX_ACTIVE_TEAMS_NORMAL_DAY = NUMBER_OF_CARS * TEAMS_PER_CAR; // 4 teams/day
const MAX_NON_VIP_ACTIVE_TEAMS_ON_VIP_DAY = (NUMBER_OF_CARS - 1) * TEAMS_PER_CAR; // 2 teams/day (VIP uses dedicated car)
// NOTE: We use location as a preference when choosing days (to group jobs by area),
// but we do NOT hard-limit locations per team per day anymore. This avoids missing visits.

/**
 * Normalize team names so small spelling / spacing differences in the JSON
 * do not create extra "teams" in the schedule.
 *
 * Canonical teams we expect:
 * - "Alemayehu & Ashebir"
 * - "Anwar & Tezazu"
 * - "Ashenafi & Aman"
 * - "Assefa & Getasew"
 */
function normalizeTeamName(raw) {
    if (!raw) return raw;

    const key = raw.toLowerCase().replace(/\s+/g, '');

    const mapping = {
        'alemayehu&ashebir': 'Alemayehu & Ashebir',
        'anwar&tezazu': 'Anwar & Tezazu',
        'ashnfi&aman': 'Ashenafi & Aman',
        'ashenafi&aman': 'Ashenafi & Aman',
        'assefa&getasew': 'Assefa & Getasew'
    };

    return mapping[key] || raw.trim();
}

class PestControlScheduler {
    constructor() {
        this.customers = [];
        this.calendar = [];
        this.schedule = new Map(); // Date string -> { teamName: [customers] }
        this.teams = new Set();
        this.vipTeams = new Set();
        this.summary = {
            totalCustomers: 0,
            totalVisits: 0,
            visitsByFrequency: {},
            visitsByTeam: {},
            violations: []
        };
    }

    /**
     * Load customers from JSON file
     */
    async loadCustomers() {
        const filePath = path.join(__dirname, 'customers.json');
        const rawData = fs.readFileSync(filePath, 'utf8');
        this.customers = JSON.parse(rawData);

        // Special Rule: EDR is the VIP customer
        this.customers.forEach(c => {
            // Normalize team names so we always end up with exactly 4 logical teams.
            c.team = normalizeTeamName(c.team);

            if (c.name === 'EDR') {
                c.frequency = 'VIP_Every_2_Days';
            }
            this.teams.add(c.team);
            this.summary.totalCustomers++;
        });

        // Track which teams have VIP routes (used for logistics constraints)
        this.customers.forEach(c => {
            if (c.frequency === 'VIP_Every_2_Days') {
                this.vipTeams.add(c.team);
            }
        });

        // Initialize schedule map for all days in March
        this.generateCalendar();
    }

    /**
     * Utility: distinct teams that have >=1 job on a day
     */
    getActiveTeams(dayData) {
        return Object.entries(dayData.jobs)
            .filter(([, jobs]) => jobs.length > 0)
            .map(([team]) => team);
    }

    /**
     * Utility: does this day already contain any VIP job?
     */
    dayHasVipJob(dayData) {
        return Object.values(dayData.jobs).some(teamJobs => teamJobs.some(j => j.type === 'VIP'));
    }

    /**
     * Utility: does a given team already have a VIP job on this day?
     */
    teamHasVipJob(dayData, team) {
        const jobs = dayData.jobs[team] || [];
        return jobs.some(j => j.type === 'VIP');
    }

    /**
     * Utility: set of distinct locations a team is visiting on a given day
     */
    getTeamLocationsForDay(dayData, team) {
        const jobs = dayData.jobs[team] || [];
        const locations = new Set();
        jobs.forEach(j => {
            if (j.location) {
                locations.add(j.location);
            }
        });
        return locations;
    }

    /**
     * Utility: whether a team already has at least one job in this customer's location on this day
     */
    teamHasSameLocation(dayData, team, customer) {
        if (!customer.location) return false;
        const locations = this.getTeamLocationsForDay(dayData, team);
        return locations.has(customer.location);
    }

    /**
     * Generate all days for March
     */
    generateCalendar() {
        const start = startOfMonth(new Date(MARCH_YEAR, MONTH_INDEX));
        const end = endOfMonth(start);
        this.calendar = eachDayOfInterval({ start, end });

        this.calendar.forEach(date => {
            const dateStr = format(date, 'yyyy-MM-dd');
            const dayData = {
                date,
                isWorkingDay: !isSunday(date),
                jobs: {}
            };
            this.teams.forEach(team => {
                dayData.jobs[team] = [];
            });
            this.schedule.set(dateStr, dayData);
        });
    }

    /**
     * Check if a job can be scheduled on a specific day for a team
     */
    canSchedule(date, team, customer, forceSunday = false) {
        if (isSunday(date) && !forceSunday) return false;

        const dateStr = format(date, 'yyyy-MM-dd');
        const dayData = this.schedule.get(dateStr);

        // Logistics: if VIP is scheduled on this date, reserve one car exclusively for VIP route.
        // That means only 2 other (non-VIP) teams can work that day (total 3 teams: VIP team + 2 other teams).
        const isVipCustomer = customer.frequency === 'VIP_Every_2_Days';
        const dayHasVip = this.dayHasVipJob(dayData);

        if (dayHasVip) {
            // If a team has the VIP route that day, it should not take non-VIP jobs (dedicated car + far location).
            if (this.teamHasVipJob(dayData, team) && !isVipCustomer) return false;

            // Non-VIP teams allowed (besides VIP team(s)) are limited to 2 teams total for the whole day.
            if (!this.vipTeams.has(team)) {
                const activeTeams = this.getActiveTeams(dayData);
                const activeNonVipTeams = activeTeams.filter(t => !this.vipTeams.has(t));
                const teamAlreadyActive = (dayData.jobs[team] || []).length > 0;

                if (!teamAlreadyActive && activeNonVipTeams.length >= MAX_NON_VIP_ACTIVE_TEAMS_ON_VIP_DAY) {
                    return false;
                }
            }
        } else {
            // Normal day: 2 cars * 2 teams per car => max 4 active teams/day (supports >4 teams if future data grows).
            const activeTeams = this.getActiveTeams(dayData);
            const teamAlreadyActive = (dayData.jobs[team] || []).length > 0;
            if (!teamAlreadyActive && activeTeams.length >= MAX_ACTIVE_TEAMS_NORMAL_DAY) return false;
        }

        // Check daily total capacity
        let dailyTotal = 0;
        Object.values(dayData.jobs).forEach(teamJobs => {
            dailyTotal += teamJobs.length;
        });

        if (dailyTotal >= CAPACITY_LIMIT) return false;

        return true;
    }

    /**
     * Add job to the schedule
     */
    addJob(date, team, customer, typeLabel) {
        const dateStr = format(date, 'yyyy-MM-dd');
        const dayData = this.schedule.get(dateStr);
        dayData.jobs[team].push({
            name: customer.name,
            frequency: customer.frequency,
            type: typeLabel, // Monthly_1, Monthly_2, etc.
            location: customer.location || ''
        });

        // Update summary
        this.summary.totalVisits++;
        this.summary.visitsByFrequency[customer.frequency] = (this.summary.visitsByFrequency[customer.frequency] || 0) + 1;
        this.summary.visitsByTeam[team] = (this.summary.visitsByTeam[team] || 0) + 1;
    }

    /**
     * Schedule VIP (Every 2 days, including Sundays)
     */
    scheduleVIP() {
        const vips = this.customers.filter(c => c.frequency === 'VIP_Every_2_Days');
        vips.forEach(customer => {
            // Start on March 1st for consistency
            let currentDate = startOfMonth(new Date(MARCH_YEAR, MONTH_INDEX));
            while (currentDate.getMonth() === MONTH_INDEX) {
                if (this.canSchedule(currentDate, customer.team, customer, true)) {
                    this.addJob(currentDate, customer.team, customer, 'VIP');
                } else {
                    this.summary.violations.push(`Could not schedule VIP for ${customer.name} on ${format(currentDate, 'yyyy-MM-dd')}`);
                }
                currentDate = addDays(currentDate, 2);
            }
        });
    }

    /**
     * Schedule Bi-Monthly
     * Business interpretation: 1 visit per month (in March).
     */
    scheduleBiMonthly() {
        const biMonthly = this.customers.filter(c => c.frequency === 'Bi-Monthly');
        biMonthly.forEach(customer => {
            // Consider all working days in March, choose the best single day.
            const candidateDays = Array.from(this.schedule.values())
                .filter(d => d.isWorkingDay)
                .sort((a, b) => {
                    const loadA = a.jobs[customer.team].length;
                    const loadB = b.jobs[customer.team].length;
                    if (loadA !== loadB) return loadA - loadB;

                    // Prefer days where this team is already in the same location (less driving).
                    const sameLocA = this.teamHasSameLocation(a, customer.team, customer);
                    const sameLocB = this.teamHasSameLocation(b, customer.team, customer);
                    if (sameLocA !== sameLocB) return sameLocA ? -1 : 1;

                    // Secondary sort by total daily load
                    return Object.values(a.jobs).flat().length - Object.values(b.jobs).flat().length;
                });

            let scheduled = false;
            for (const dayData of candidateDays) {
                if (this.canSchedule(dayData.date, customer.team, customer)) {
                    this.addJob(dayData.date, customer.team, customer, 'Bi-Monthly');
                    scheduled = true;
                    break;
                }
            }

            if (!scheduled) {
                this.summary.violations.push(`Could not schedule Bi-Monthly visit for ${customer.name}`);
            }
        });
    }

    /**
     * Schedule Monthly (2 visits, 12-16 days apart)
     * Spreads 1st visits across the first half to ensure the 2nd visits fill the second half.
     */
    scheduleMonthly() {
        const monthly = this.customers.filter(c => c.frequency === 'Monthly');

        // Randomize order slightly to prevent team clustering
        monthly.sort(() => Math.random() - 0.5);

        monthly.forEach(customer => {
            let firstVisitDate = null;
            let secondVisitDate = null;

            // Find best day for 1st visit (1-15) based on team load
            const firstHalfDays = Array.from(this.schedule.values())
                .filter(d => d.date.getDate() <= 15 && d.isWorkingDay)
                .sort((a, b) => {
                    const loadA = a.jobs[customer.team].length;
                    const loadB = b.jobs[customer.team].length;
                    if (loadA !== loadB) return loadA - loadB;

                    const sameLocA = this.teamHasSameLocation(a, customer.team, customer);
                    const sameLocB = this.teamHasSameLocation(b, customer.team, customer);
                    if (sameLocA !== sameLocB) return sameLocA ? -1 : 1;

                    return Object.values(a.jobs).flat().length - Object.values(b.jobs).flat().length;
                });

            for (const dayData of firstHalfDays) {
                if (this.canSchedule(dayData.date, customer.team, customer)) {
                    this.addJob(dayData.date, customer.team, customer, 'Monthly_1');
                    firstVisitDate = dayData.date;
                    break;
                }
            }

            if (!firstVisitDate) {
                this.summary.violations.push(`Missing Monthly_1 for ${customer.name}`);
                return;
            }

            // Find best day for 2nd visit (12-16 days later) based on team load
            const secondHalfDays = Array.from(this.schedule.values())
                .filter(d => {
                    const gap = differenceInDays(d.date, firstVisitDate);
                    return gap >= 12 && gap <= 16 && d.isWorkingDay;
                })
                .sort((a, b) => {
                    const loadA = a.jobs[customer.team].length;
                    const loadB = b.jobs[customer.team].length;
                    if (loadA !== loadB) return loadA - loadB;

                    const sameLocA = this.teamHasSameLocation(a, customer.team, customer);
                    const sameLocB = this.teamHasSameLocation(b, customer.team, customer);
                    if (sameLocA !== sameLocB) return sameLocA ? -1 : 1;

                    return Object.values(a.jobs).flat().length - Object.values(b.jobs).flat().length;
                });

            for (const dayData of secondHalfDays) {
                if (this.canSchedule(dayData.date, customer.team, customer)) {
                    this.addJob(dayData.date, customer.team, customer, 'Monthly_2');
                    secondVisitDate = dayData.date;
                    break;
                }
            }

            if (!secondVisitDate) {
                this.summary.violations.push(`Missing Monthly_2 for ${customer.name}`);
            }
        });
    }

    /**
     * Schedule Quarterly (1 visit in March)
     * Fills the biggest gaps for each team.
     */
    scheduleQuarterly() {
        const quarterly = this.customers.filter(c => c.frequency === 'Quarterly');

        quarterly.forEach(customer => {
            let scheduled = false;

            // Prioritize days where the team is currently idle or has lowest load
            const candidateDays = Array.from(this.schedule.values())
                .filter(d => d.isWorkingDay)
                .sort((a, b) => {
                    const teamLoadA = a.jobs[customer.team].length;
                    const teamLoadB = b.jobs[customer.team].length;
                    if (teamLoadA !== teamLoadB) return teamLoadA - teamLoadB;
                    const sameLocA = this.teamHasSameLocation(a, customer.team, customer);
                    const sameLocB = this.teamHasSameLocation(b, customer.team, customer);
                    if (sameLocA !== sameLocB) return sameLocA ? -1 : 1;
                    return Object.values(a.jobs).flat().length - Object.values(b.jobs).flat().length;
                });

            for (const dayData of candidateDays) {
                if (this.canSchedule(dayData.date, customer.team, customer)) {
                    this.addJob(dayData.date, customer.team, customer, 'Quarterly');
                    scheduled = true;
                    break;
                }
            }

            if (!scheduled) {
                this.summary.violations.push(`Could not schedule Quarterly for ${customer.name}`);
            }
        });
    }

    /**
     * Validate final schedule against rules
     */
    validateSchedule() {
        this.schedule.forEach(day => {
            const total = Object.values(day.jobs).flat().length;
            if (total > CAPACITY_LIMIT && !isSunday(day.date)) {
                // Technically VIP can happen on Sunday, but for working days we target 18
            }

            // Logistics validation:
            // - Sundays: only VIP jobs should exist
            // - VIP days: at most 2 non-VIP teams can have jobs
            const dayHasVip = this.dayHasVipJob(day);
            const activeTeams = this.getActiveTeams(day);
            const activeNonVipTeams = activeTeams.filter(t => !this.vipTeams.has(t));

            if (isSunday(day.date)) {
                const nonVipJobsExist = Object.entries(day.jobs).some(([team, jobs]) => {
                    if (this.vipTeams.has(team)) return false;
                    return jobs.length > 0;
                });
                if (nonVipJobsExist) {
                    this.summary.violations.push(`Sunday logistics violation on ${format(day.date, 'yyyy-MM-dd')}: non-VIP teams scheduled.`);
                }
            }

            if (dayHasVip && activeNonVipTeams.length > MAX_NON_VIP_ACTIVE_TEAMS_ON_VIP_DAY) {
                this.summary.violations.push(
                    `VIP-day logistics violation on ${format(day.date, 'yyyy-MM-dd')}: ${activeNonVipTeams.length} non-VIP teams scheduled (max ${MAX_NON_VIP_ACTIVE_TEAMS_ON_VIP_DAY}).`
                );
            }

            // VIP team must be dedicated: if VIP job exists for a VIP team, it can't also have non-VIP jobs that day
            this.vipTeams.forEach(vipTeam => {
                const jobs = day.jobs[vipTeam] || [];
                const hasVip = jobs.some(j => j.type === 'VIP');
                if (!hasVip) return;
                const hasNonVip = jobs.some(j => j.type !== 'VIP');
                if (hasNonVip) {
                    this.summary.violations.push(
                        `VIP-team dedication violation on ${format(day.date, 'yyyy-MM-dd')}: team "${vipTeam}" has VIP + non-VIP jobs same day.`
                    );
                }
            });
        });

        // Check if any customer missed their expected visit count
        this.customers.forEach(customer => {
            const jobs = [];
            this.schedule.forEach(day => {
                const found = day.jobs[customer.team].find(j => j.name === customer.name);
                if (found) jobs.push({ date: day.date, ...found });
            });

            if (customer.frequency === 'Monthly' && jobs.length !== 2) {
                // Already flagged in scheduleMonthly
            }

            if (customer.frequency === 'Monthly' && jobs.length === 2) {
                const gap = differenceInDays(jobs[1].date, jobs[0].date);
                if (gap < 12 || gap > 16) {
                    this.summary.violations.push(`Monthly violation for ${customer.name}: gap is ${gap} days.`);
                }
            }

            if (customer.frequency === 'Bi-Monthly' && jobs.length !== 1) {
                this.summary.violations.push(`Bi-Monthly violation for ${customer.name}: expected 1 visit, got ${jobs.length}.`);
            }
        });

        return this.summary.violations.length === 0;
    }

    /**
     * Export the schedule to Excel with formatting
     */
    async exportToExcel() {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('March Schedule');

        // Define columns
        const teamNames = Array.from(this.teams).sort();
        const columns = [
            { header: 'DATE', key: 'date', width: 14 },
            { header: 'DAY', key: 'day', width: 12 },
            ...teamNames.map((name, i) => ({ header: `TEAM ${i + 1}\n${name.toUpperCase()}`, key: `team_${i}`, width: 38 })),
            { header: 'TOTAL', key: 'total', width: 10 }
        ];
        sheet.columns = columns;

        // Apply header styling (Row 1)
        const headerRow = sheet.getRow(1);
        headerRow.height = 35;
        headerRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 9, name: 'Segoe UI' };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '333333' } // Dark Charcoal
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                bottom: { style: 'medium', color: { argb: '000000' } }
            };
        });

        // Add Data Rows
        this.schedule.forEach((dayData) => {
            const rowData = {
                date: format(dayData.date, 'MMM dd, yyyy'),
                day: format(dayData.date, 'EEEE').toUpperCase(),
                total: 0
            };

            teamNames.forEach((team, idx) => {
                const jobs = dayData.jobs[team];
                // Placeholder; we will overwrite with rich text after adding the row
                rowData[`team_${idx}`] = jobs.length ? ' ' : '';
                rowData.total += jobs.length;
            });

            const row = sheet.addRow(rowData);
            row.height = 65; // Balanced height

            // Default cell style
            row.eachCell((cell) => {
                cell.font = { name: 'Segoe UI', size: 8.5 }; // Smaller font for elegance
                cell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left', padding: { left: 4, top: 4 } };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'F2F2F2' } },
                    left: { style: 'thin', color: { argb: 'F2F2F2' } },
                    bottom: { style: 'thin', color: { argb: 'F2F2F2' } },
                    right: { style: 'thin', color: { argb: 'F2F2F2' } }
                };
            });

            // Center basic info columns
            row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
            row.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };
            row.getCell(sheet.columns.length).alignment = { vertical: 'middle', horizontal: 'center' };

            // Apply Coloring Rules and rich text with faded location
            teamNames.forEach((team, idx) => {
                const jobs = dayData.jobs[team];
                const cell = row.getCell(3 + idx);

                if (jobs.length > 0) {
                    const type = jobs[0].type;
                    let bgColor = 'FFFFFF';
                    let mainTextColor = '000000';

                    // Refined, softer color palette
                    if (type === 'VIP') {
                        bgColor = 'FFD9D9'; // Soft Red
                        mainTextColor = '990000'; // Dark Red Text
                    } else if (type === 'Bi-Monthly') {
                        bgColor = 'FFF2CC'; // Soft Orange/Yellow
                        mainTextColor = '996600'; // Brownish Text
                    } else if (type === 'Monthly_1') {
                        bgColor = 'D9EAD3'; // Soft Green
                        mainTextColor = '274E13'; // Dark Green Text
                    } else if (type === 'Monthly_2') {
                        bgColor = 'D0E2F3'; // Soft Blue
                        mainTextColor = '0B5394'; // Dark Blue Text
                    } else if (type === 'Quarterly') {
                        bgColor = 'EAD1DC'; // Soft Purple/Pink
                        mainTextColor = '741B47'; // Dark Purple Text
                    }

                    // Background color per frequency
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: bgColor }
                    };

                    // Build rich text: "• Name (Location)" with location in faded gray
                    const richText = [];
                    jobs.forEach((j, i) => {
                        const nameText = `• ${j.name}`;
                        const locText = j.location ? ` (${j.location})` : '';

                        richText.push({
                            text: nameText,
                            font: {
                                color: { argb: mainTextColor },
                                size: 8.5,
                                name: 'Segoe UI',
                                bold: true
                            }
                        });

                        if (locText) {
                            richText.push({
                                text: locText,
                                font: {
                                    color: { argb: '999999' }, // faded gray
                                    size: 8.0,
                                    name: 'Segoe UI',
                                    italic: true
                                }
                            });
                        }

                        if (i < jobs.length - 1) {
                            richText.push({ text: '\n' });
                        }
                    });

                    cell.value = { richText };
                }
            });

            // Daily Total styling
            const totalCell = row.getCell(sheet.columns.length);
            totalCell.font = { bold: true, name: 'Segoe UI', size: 9 };
            if (rowData.total > CAPACITY_LIMIT) {
                totalCell.font = { color: { argb: 'FF0000' }, bold: true, size: 10 };
            }

            // Shade Sundays
            if (isSunday(dayData.date)) {
                row.eachCell((cell) => {
                    const currentFill = (cell.fill && cell.fill.fgColor) ? cell.fill.fgColor.argb : 'FFFFFF';
                    if (currentFill === 'FFFFFF') {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FAFAFA' }
                        };
                        cell.font = { color: { argb: 'CCCCCC' }, italic: true, name: 'Segoe UI', size: 8 };
                    }
                });
            }
        });

        // Summary Report Sheet
        const summarySheet = workbook.addWorksheet('Summary Report');
        summarySheet.columns = [
            { header: 'Metric', key: 'metric', width: 30 },
            { header: 'Value', key: 'value', width: 30 }
        ];

        const avgJobs = (this.summary.totalVisits / this.calendar.filter(d => !isSunday(d)).length).toFixed(2);

        // Find peak load
        let peakLoad = 0;
        let peakDay = '';
        this.schedule.forEach(day => {
            const total = Object.values(day.jobs).flat().length;
            if (total > peakLoad) {
                peakLoad = total;
                peakDay = format(day.date, 'yyyy-MM-dd');
            }
        });

        summarySheet.addRows([
            ['Total Customers', this.summary.totalCustomers],
            ['Total Visits Scheduled', this.summary.totalVisits],
            ['', ''],
            ['--- VISITS PER FREQUENCY ---', ''],
            ...Object.entries(this.summary.visitsByFrequency),
            ['', ''],
            ['--- VISITS PER TEAM ---', ''],
            ...Object.entries(this.summary.visitsByTeam),
            ['', ''],
            ['Average jobs per working day', avgJobs],
            ['Peak load day', `${peakDay} (${peakLoad} jobs)`],
            ['Validation Check', this.summary.violations.length === 0 ? 'PASSED' : 'FAILED'],
        ]);

        if (this.summary.violations.length > 0) {
            summarySheet.addRow(['', '']);
            summarySheet.addRow(['--- VIOLATIONS ---', '']);
            this.summary.violations.forEach(v => summarySheet.addRow([v, '']));
        }

        const outputPath = path.join(__dirname, 'March_Pest_Control_Schedule.xlsx');
        await workbook.xlsx.writeFile(outputPath);
        return outputPath;
    }

    async run() {
        console.log('--- Starting Pest Control Scheduler ---');

        try {
            await this.loadCustomers();
            console.log(`Loaded ${this.customers.length} customers.`);

            // Priority: VIP -> Bi-Monthly -> Monthly -> Quarterly
            this.scheduleVIP();
            this.scheduleBiMonthly();
            this.scheduleMonthly();
            this.scheduleQuarterly();

            const isValid = this.validateSchedule();
            const outputPath = await this.exportToExcel();

            console.log('\n--- Generation Summary ---');
            console.log(`Total Visits: ${this.summary.totalVisits}`);
            console.log(`Validation: ${isValid ? 'PASSED' : 'FAILED'}`);
            if (!isValid) {
                console.log(`Violations Found: ${this.summary.violations.length}`);
                this.summary.violations.forEach(v => console.log(` - ${v}`));
            }
            console.log(`\nVisits by Team:`, this.summary.visitsByTeam);
            console.log(`Visits by Frequency:`, this.summary.visitsByFrequency);
            console.log(`\nExcel file generated: ${outputPath}`);

        } catch (error) {
            console.error('CRITICAL ERROR:', error.message);
            process.exit(1);
        }
    }
}

// Execute
const scheduler = new PestControlScheduler();
scheduler.run();
