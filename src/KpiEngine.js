/**
 * Project Management Workspace - Presence & Remote Work KPI Engine
 * Ticket 06: Presence & Remote Work KPI Engine (calc_kpis)
 *
 * Computes monthly aggregated attendance metrics across members and departments:
 * - Weights AM and PM periods at 0.5 working days each for intra-day mobility.
 * - Computes Remote rates and In-Office rates.
 * - Provides office-specific occupancy breakdowns (Main Office, Downtown Office, Branch Office, Client Site).
 * - Incorporates FTE rates from tbl_tenures to evaluate contract capacity vs actual attendance.
 * - Generates formatted calc_kpis dashboard with 12px "Google Sans" and native percentage formatting.
 */

/**
 * Normalizes and parses raw date cell into Date object at midnight.
 *
 * @param {Date|string|number} rawDate
 * @return {Date|null}
 */
function parseKpiDate(rawDate) {
  if (!rawDate) return null;
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    var d = new Date(rawDate.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  var str = String(rawDate).trim();
  var isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10), 0, 0, 0, 0);
  }
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    parsed.setHours(0, 0, 0, 0);
    return parsed;
  }
  return null;
}

/**
 * Returns month key in YYYY-MM format.
 *
 * @param {Date} date
 * @return {string}
 */
function formatMonthKey(date) {
  if (!date || isNaN(date.getTime())) return "";
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1);
  if (m.length < 2) m = "0" + m;
  return y + "-" + m;
}

/**
 * Returns total number of weekdays (Mon-Fri) in a given YYYY-MM month.
 *
 * @param {string} monthKey - E.g. "2026-01"
 * @return {number} Count of weekdays
 */
function countMonthWeekdays(monthKey) {
  if (!monthKey || typeof monthKey !== "string") return 0;
  var parts = monthKey.split("-");
  if (parts.length !== 2) return 0;
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10); // 1-12
  if (isNaN(year) || isNaN(month)) return 0;

  var startDate = new Date(year, month - 1, 1, 0, 0, 0, 0);
  var endDate = new Date(year, month, 0, 0, 0, 0, 0); // Last day of month

  var count = 0;
  var curr = new Date(startDate.getTime());
  while (curr <= endDate) {
    var day = curr.getDay();
    if (day >= 1 && day <= 5) {
      count++;
    }
    curr.setDate(curr.getDate() + 1);
  }
  return count;
}

/**
 * Determines whether a location string represents a non-working leave or holiday.
 *
 * @param {string} location
 * @return {boolean}
 */
function isLeaveLocation(location) {
  if (!location) return false;
  var loc = String(location).trim().toLowerCase();
  return (
    loc.indexOf("vacation") !== -1 ||
    loc.indexOf("sick") !== -1 ||
    loc.indexOf("holiday") !== -1 ||
    loc.indexOf("unpaid") !== -1 ||
    loc.indexOf("maternity") !== -1 ||
    loc.indexOf("paternity") !== -1 ||
    loc === "leave" ||
    loc === "off" ||
    loc === "absence"
  );
}

/**
 * Categorizes a work location into one of standard buckets:
 * "remote", "main_office", "downtown_office", "branch_office", "client_site", "other_office", or "none" (if leave or empty).
 *
 * @param {string} location
 * @return {string}
 */
function categorizeLocation(location) {
  if (!location) return "none";
  var loc = String(location).trim();
  if (!loc) return "none";
  if (isLeaveLocation(loc)) return "none";

  var lower = loc.toLowerCase();
  if (lower === "remote" || lower.indexOf("home") !== -1) {
    return "remote";
  }
  if (lower.indexOf("main office") !== -1) {
    return "main_office";
  }
  if (lower.indexOf("downtown") !== -1) {
    return "downtown_office";
  }
  if (lower.indexOf("branch") !== -1) {
    return "branch_office";
  }
  if (lower.indexOf("client") !== -1) {
    return "client_site";
  }
  return "other_office";
}

/**
 * Computes all presence, remote work, occupancy, and capacity KPIs from operational tables.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Structured KPI calculation model
 */
function computePresenceKpis(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet available" };

  var membersSheet = ss.getSheetByName("tbl_members");
  var tenuresSheet = ss.getSheetByName("tbl_tenures");
  var logSheet = ss.getSheetByName("tbl_presence_log");

  if (!logSheet) {
    return { success: false, message: "Sheet 'tbl_presence_log' not found" };
  }

  // 1. Build members roster from tbl_members
  var membersRoster = {};
  if (membersSheet && membersSheet.getMaxRows() > 1) {
    var mMax = Math.min(membersSheet.getMaxRows(), 500);
    var mRange = membersSheet.getRange(2, 1, mMax - 1, Math.max(membersSheet.getLastColumn(), 7));
    var mVals = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];

    for (var m = 0; m < mVals.length; m++) {
      var rawName = mVals[m][0];
      var richName = (mRich && mRich[m]) ? mRich[m][0] : null;
      var nameStr = (typeof extractCellValue === "function") ? extractCellValue(rawName, richName) : String(rawName || "").trim();
      if (!nameStr) continue;

      var email = String(mVals[m][1] || "").trim().toLowerCase();
      var dept = String(mVals[m][4] || "").trim();
      var status = String(mVals[m][5] || "Active").trim();

      var mData = {
        name: nameStr,
        email: email,
        department: dept || "General",
        status: status,
        sourceRow: m + 2
      };

      membersRoster[nameStr.toLowerCase()] = mData;
      if (email) membersRoster[email] = mData;
    }
  }

  // 2. Build tenures records from tbl_tenures
  var tenuresList = [];
  if (tenuresSheet && tenuresSheet.getMaxRows() > 1) {
    var tMax = Math.min(tenuresSheet.getMaxRows(), 500);
    var tRange = tenuresSheet.getRange(2, 1, tMax - 1, Math.max(tenuresSheet.getLastColumn(), 7));
    var tVals = tRange.getValues();
    var tRich = (typeof tRange.getRichTextValues === "function") ? tRange.getRichTextValues() : [];

    for (var t = 0; t < tVals.length; t++) {
      var tRaw = tVals[t][0];
      var tR = (tRich && tRich[t]) ? tRich[t][0] : null;
      var tName = (typeof extractCellValue === "function") ? extractCellValue(tRaw, tR) : String(tRaw || "").trim();
      if (!tName) continue;

      var dept = String(tVals[t][2] || "").trim();
      var sDate = parseKpiDate(tVals[t][3]);
      var eDate = parseKpiDate(tVals[t][4]);
      var fteRaw = tVals[t][6];
      var fte = (fteRaw !== "" && !isNaN(parseFloat(fteRaw))) ? parseFloat(fteRaw) : 1.0;

      tenuresList.push({
        name: tName,
        department: dept,
        startDate: sDate,
        endDate: eDate,
        fteRate: fte
      });
    }
  }

  // Helper to resolve department and FTE for a member in a specific month
  function resolveMemberTenure(memberName, monthKey) {
    var mParts = monthKey.split("-");
    var mYear = parseInt(mParts[0], 10);
    var mNum = parseInt(mParts[1], 10);
    var mStart = new Date(mYear, mNum - 1, 1, 0, 0, 0, 0);
    var mEnd = new Date(mYear, mNum, 0, 23, 59, 59, 999);

    var memberKey = memberName.toLowerCase();
    var rosterMember = membersRoster[memberKey] || null;

    var matchedTenure = null;
    for (var i = 0; i < tenuresList.length; i++) {
      var ten = tenuresList[i];
      var matchesName = false;
      if (ten.name.toLowerCase() === memberKey) {
        matchesName = true;
      } else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(ten.name, null, memberName)) {
        matchesName = true;
      }

      if (matchesName) {
        var startOk = (!ten.startDate || ten.startDate <= mEnd);
        var endOk = (!ten.endDate || ten.endDate >= mStart);
        if (startOk && endOk) {
          matchedTenure = ten;
          break;
        }
      }
    }

    var dept = "";
    var fte = 1.0;

    if (matchedTenure) {
      dept = matchedTenure.department;
      fte = matchedTenure.fteRate;
    }
    if (!dept && rosterMember) {
      dept = rosterMember.department;
    }
    if (!dept) dept = "General";

    return {
      department: dept,
      fteRate: fte,
      rosterMember: rosterMember
    };
  }

  // 3. Process logs from tbl_presence_log
  var lMax = Math.min(logSheet.getMaxRows(), 5000);
  var memberMonthMetrics = {}; // Key: monthKey + "::" + memberName
  var distinctMonths = {};

  if (lMax > 1) {
    var lRange = logSheet.getRange(2, 1, lMax - 1, Math.max(logSheet.getLastColumn(), 6));
    var lVals = lRange.getValues();
    var lRich = (typeof lRange.getRichTextValues === "function") ? lRange.getRichTextValues() : [];

    for (var r = 0; r < lVals.length; r++) {
      var dateObj = parseKpiDate(lVals[r][1]);
      if (!dateObj) continue;

      var rawMem = lVals[r][2];
      var richMem = (lRich && lRich[r]) ? lRich[r][2] : null;
      var memName = (typeof extractCellValue === "function") ? extractCellValue(rawMem, richMem) : String(rawMem || "").trim();
      if (!memName) continue;

      var monthKey = formatMonthKey(dateObj);
      if (!monthKey) continue;
      distinctMonths[monthKey] = true;

      var key = monthKey + "::" + memName;
      if (!memberMonthMetrics[key]) {
        memberMonthMetrics[key] = {
          month: monthKey,
          member: memName,
          amSlots: 0,
          pmSlots: 0,
          workedPeriods: 0,
          leavePeriods: 0,
          remotePeriods: 0,
          inOfficePeriods: 0,
          mainOfficePeriods: 0,
          downtownPeriods: 0,
          branchPeriods: 0,
          clientSitePeriods: 0
        };
      }

      var entry = memberMonthMetrics[key];
      var locAm = lVals[r][3];
      var locPm = lVals[r][4];

      // Evaluate AM period
      if (locAm) {
        entry.amSlots++;
        if (isLeaveLocation(locAm)) {
          entry.leavePeriods++;
        } else {
          var catAm = categorizeLocation(locAm);
          if (catAm !== "none") {
            entry.workedPeriods++;
            if (catAm === "remote") {
              entry.remotePeriods++;
            } else {
              entry.inOfficePeriods++;
              if (catAm === "main_office") entry.mainOfficePeriods++;
              else if (catAm === "downtown_office") entry.downtownPeriods++;
              else if (catAm === "branch_office") entry.branchPeriods++;
              else if (catAm === "client_site") entry.clientSitePeriods++;
            }
          }
        }
      }

      // Evaluate PM period
      if (locPm) {
        entry.pmSlots++;
        if (isLeaveLocation(locPm)) {
          entry.leavePeriods++;
        } else {
          var catPm = categorizeLocation(locPm);
          if (catPm !== "none") {
            entry.workedPeriods++;
            if (catPm === "remote") {
              entry.remotePeriods++;
            } else {
              entry.inOfficePeriods++;
              if (catPm === "main_office") entry.mainOfficePeriods++;
              else if (catPm === "downtown_office") entry.downtownPeriods++;
              else if (catPm === "branch_office") entry.branchPeriods++;
              else if (catPm === "client_site") entry.clientSitePeriods++;
            }
          }
        }
      }
    }
  }

  // If no months found, default to current month
  var sortedMonths = Object.keys(distinctMonths).sort().reverse();
  if (sortedMonths.length === 0) {
    sortedMonths.push(formatMonthKey(new Date()));
  }

  // Also ensure members from tbl_members exist in each distinct month
  for (var mIdx = 0; mIdx < sortedMonths.length; mIdx++) {
    var mon = sortedMonths[mIdx];
    for (var k in membersRoster) {
      if (!membersRoster.hasOwnProperty(k)) continue;
      var mb = membersRoster[k];
      if (mb.status === "Departed" || mb.status === "Archived") continue;
      var mKey = mon + "::" + mb.name;
      if (!memberMonthMetrics[mKey]) {
        memberMonthMetrics[mKey] = {
          month: mon,
          member: mb.name,
          amSlots: 0,
          pmSlots: 0,
          workedPeriods: 0,
          leavePeriods: 0,
          remotePeriods: 0,
          inOfficePeriods: 0,
          mainOfficePeriods: 0,
          downtownPeriods: 0,
          branchPeriods: 0,
          clientSitePeriods: 0
        };
      }
    }
  }

  // 4. Transform raw period counts into calculated analytical KPIs
  var calculatedRows = [];
  var departmentAggregates = {}; // Key: month + "::" + department
  var monthlyTotals = {};        // Key: month

  for (var rowKey in memberMonthMetrics) {
    if (!memberMonthMetrics.hasOwnProperty(rowKey)) continue;
    var rowData = memberMonthMetrics[rowKey];

    var tenureInfo = resolveMemberTenure(rowData.member, rowData.month);
    var dept = tenureInfo.department;
    var fte = tenureInfo.fteRate;
    var rosterObj = tenureInfo.rosterMember;

    var monthWeekdays = countMonthWeekdays(rowData.month);
    var contractCapacity = monthWeekdays * fte;

    var workedDays = rowData.workedPeriods * 0.5;
    var leaveDays = rowData.leavePeriods * 0.5;
    var capacityUtilization = (contractCapacity > 0) ? (workedDays / contractCapacity) : 0;

    var remoteDays = rowData.remotePeriods * 0.5;
    var remoteRate = (workedDays > 0) ? (remoteDays / workedDays) : 0;

    var inOfficeDays = rowData.inOfficePeriods * 0.5;
    var inOfficeRate = (workedDays > 0) ? (inOfficeDays / workedDays) : 0;

    var mainOfficeDays = rowData.mainOfficePeriods * 0.5;
    var downtownDays = rowData.downtownPeriods * 0.5;
    var branchDays = rowData.branchPeriods * 0.5;
    var clientSiteDays = rowData.clientSitePeriods * 0.5;

    var calculated = {
      month: rowData.month,
      department: dept,
      member: rowData.member,
      rosterSourceRow: rosterObj ? rosterObj.sourceRow : -1,
      fteRate: fte,
      monthWeekdays: monthWeekdays,
      contractCapacityDays: contractCapacity,
      workedDays: workedDays,
      leaveDays: leaveDays,
      capacityUtilization: capacityUtilization,
      remoteDays: remoteDays,
      remoteRate: remoteRate,
      inOfficeDays: inOfficeDays,
      inOfficeRate: inOfficeRate,
      mainOfficeDays: mainOfficeDays,
      downtownOfficeDays: downtownDays,
      branchOfficeDays: branchDays,
      clientSiteDays: clientSiteDays
    };

    calculatedRows.push(calculated);

    // Roll up into Department Aggregates
    var dKey = rowData.month + "::" + dept;
    if (!departmentAggregates[dKey]) {
      departmentAggregates[dKey] = {
        month: rowData.month,
        department: dept,
        membersCount: 0,
        totalFte: 0,
        totalContractCapacity: 0,
        totalWorkedDays: 0,
        totalLeaveDays: 0,
        totalRemoteDays: 0,
        totalInOfficeDays: 0,
        totalMainOfficeDays: 0,
        totalDowntownDays: 0,
        totalBranchDays: 0,
        totalClientSiteDays: 0
      };
    }
    var dAgg = departmentAggregates[dKey];
    dAgg.membersCount++;
    dAgg.totalFte += fte;
    dAgg.totalContractCapacity += contractCapacity;
    dAgg.totalWorkedDays += workedDays;
    dAgg.totalLeaveDays += leaveDays;
    dAgg.totalRemoteDays += remoteDays;
    dAgg.totalInOfficeDays += inOfficeDays;
    dAgg.totalMainOfficeDays += mainOfficeDays;
    dAgg.totalDowntownDays += downtownDays;
    dAgg.totalBranchDays += branchDays;
    dAgg.totalClientSiteDays += clientSiteDays;

    // Roll up into Monthly Totals
    var mKeyMon = rowData.month;
    if (!monthlyTotals[mKeyMon]) {
      monthlyTotals[mKeyMon] = {
        month: rowData.month,
        membersCount: 0,
        totalFte: 0,
        totalContractCapacity: 0,
        totalWorkedDays: 0,
        totalLeaveDays: 0,
        totalRemoteDays: 0,
        totalInOfficeDays: 0,
        totalMainOfficeDays: 0,
        totalDowntownDays: 0,
        totalBranchDays: 0,
        totalClientSiteDays: 0
      };
    }
    var mAgg = monthlyTotals[mKeyMon];
    mAgg.membersCount++;
    mAgg.totalFte += fte;
    mAgg.totalContractCapacity += contractCapacity;
    mAgg.totalWorkedDays += workedDays;
    mAgg.totalLeaveDays += leaveDays;
    mAgg.totalRemoteDays += remoteDays;
    mAgg.totalInOfficeDays += inOfficeDays;
    mAgg.totalMainOfficeDays += mainOfficeDays;
    mAgg.totalDowntownDays += downtownDays;
    mAgg.totalBranchDays += branchDays;
    mAgg.totalClientSiteDays += clientSiteDays;
  }

  // Sort member rows: Month (descending) -> Department (A-Z) -> Member (A-Z)
  calculatedRows.sort(function(a, b) {
    if (a.month !== b.month) return b.month.localeCompare(a.month);
    if (a.department !== b.department) return a.department.localeCompare(b.department);
    return a.member.localeCompare(b.member);
  });

  return {
    success: true,
    memberRows: calculatedRows,
    departmentAggregates: departmentAggregates,
    monthlyTotals: monthlyTotals,
    monthsCount: sortedMonths.length,
    distinctMonths: sortedMonths
  };
}

/**
 * Renders the calc_kpis dashboard tab in Google Sheets with complete formatting,
 * People Chips, percentage formats, and department/month rollups.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function renderKpiDashboard(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet available" };

  var dataModel = computePresenceKpis(ss);
  if (!dataModel.success) return dataModel;

  var membersSheet = ss.getSheetByName("tbl_members");

  // Locate or insert calc_kpis sheet
  var kpiSheet = ss.getSheetByName("calc_kpis");
  if (!kpiSheet) {
    kpiSheet = ss.insertSheet("calc_kpis");
  }

  var headers = [
    "Month",
    "Department",
    "Member",
    "FTE Rate",
    "Contract Capacity (Days)",
    "Worked Days",
    "Leave Days",
    "Capacity Utilization",
    "Remote Days",
    "Remote Rate",
    "In-Office Days",
    "In-Office Rate",
    "Main Office (Days)",
    "Downtown Office (Days)",
    "Branch Office (Days)",
    "Client Site (Days)"
  ];
  var numCols = headers.length;

  var memberRows = dataModel.memberRows;
  var tableValues = [];
  var rowTypeTracker = []; // "member", "dept_total", "month_total"
  var chipCopyMapping = []; // Tracks rows where People Chip should be copied

  var currentMonth = null;
  var currentDept = null;

  for (var i = 0; i < memberRows.length; i++) {
    var row = memberRows[i];

    // If month changed, or department changed, insert department subtotal for previous department if exists
    if (currentMonth !== null && (row.month !== currentMonth || row.department !== currentDept)) {
      var prevDeptKey = currentMonth + "::" + currentDept;
      var dAgg = dataModel.departmentAggregates[prevDeptKey];
      if (dAgg && dAgg.membersCount > 1) {
        var dCapUtil = (dAgg.totalContractCapacity > 0) ? (dAgg.totalWorkedDays / dAgg.totalContractCapacity) : 0;
        var dRemRate = (dAgg.totalWorkedDays > 0) ? (dAgg.totalRemoteDays / dAgg.totalWorkedDays) : 0;
        var dOffRate = (dAgg.totalWorkedDays > 0) ? (dAgg.totalInOfficeDays / dAgg.totalWorkedDays) : 0;

        tableValues.push([
          currentMonth,
          currentDept,
          "Total " + currentDept,
          dAgg.totalFte,
          dAgg.totalContractCapacity,
          dAgg.totalWorkedDays,
          dAgg.totalLeaveDays,
          dCapUtil,
          dAgg.totalRemoteDays,
          dRemRate,
          dAgg.totalInOfficeDays,
          dOffRate,
          dAgg.totalMainOfficeDays,
          dAgg.totalDowntownDays,
          dAgg.totalBranchDays,
          dAgg.totalClientSiteDays
        ]);
        rowTypeTracker.push("dept_total");
        chipCopyMapping.push(null);
      }
    }

    // If month changed, insert month total for previous month
    if (currentMonth !== null && row.month !== currentMonth) {
      var mAgg = dataModel.monthlyTotals[currentMonth];
      if (mAgg) {
        var mCapUtil = (mAgg.totalContractCapacity > 0) ? (mAgg.totalWorkedDays / mAgg.totalContractCapacity) : 0;
        var mRemRate = (mAgg.totalWorkedDays > 0) ? (mAgg.totalRemoteDays / mAgg.totalWorkedDays) : 0;
        var mOffRate = (mAgg.totalWorkedDays > 0) ? (mAgg.totalInOfficeDays / mAgg.totalWorkedDays) : 0;

        tableValues.push([
          currentMonth,
          "All Departments",
          "Total Month (" + currentMonth + ")",
          mAgg.totalFte,
          mAgg.totalContractCapacity,
          mAgg.totalWorkedDays,
          mAgg.totalLeaveDays,
          mCapUtil,
          mAgg.totalRemoteDays,
          mRemRate,
          mAgg.totalInOfficeDays,
          mOffRate,
          mAgg.totalMainOfficeDays,
          mAgg.totalDowntownDays,
          mAgg.totalBranchDays,
          mAgg.totalClientSiteDays
        ]);
        rowTypeTracker.push("month_total");
        chipCopyMapping.push(null);
      }
    }

    currentMonth = row.month;
    currentDept = row.department;

    // Add individual member row
    tableValues.push([
      row.month,
      row.department,
      row.member,
      row.fteRate,
      row.contractCapacityDays,
      row.workedDays,
      row.leaveDays,
      row.capacityUtilization,
      row.remoteDays,
      row.remoteRate,
      row.inOfficeDays,
      row.inOfficeRate,
      row.mainOfficeDays,
      row.downtownOfficeDays,
      row.branchOfficeDays,
      row.clientSiteDays
    ]);
    rowTypeTracker.push("member");
    chipCopyMapping.push({
      sourceRow: row.rosterSourceRow,
      name: row.member
    });
  }

  // Handle final department total and month total
  if (currentMonth !== null) {
    var finalDeptKey = currentMonth + "::" + currentDept;
    var finalDAgg = dataModel.departmentAggregates[finalDeptKey];
    if (finalDAgg && finalDAgg.membersCount > 1) {
      var fDCapUtil = (finalDAgg.totalContractCapacity > 0) ? (finalDAgg.totalWorkedDays / finalDAgg.totalContractCapacity) : 0;
      var fDRemRate = (finalDAgg.totalWorkedDays > 0) ? (finalDAgg.totalRemoteDays / finalDAgg.totalWorkedDays) : 0;
      var fDOffRate = (finalDAgg.totalWorkedDays > 0) ? (finalDAgg.totalInOfficeDays / finalDAgg.totalWorkedDays) : 0;

      tableValues.push([
        currentMonth,
        currentDept,
        "Total " + currentDept,
        finalDAgg.totalFte,
        finalDAgg.totalContractCapacity,
        finalDAgg.totalWorkedDays,
        finalDAgg.totalLeaveDays,
        fDCapUtil,
        finalDAgg.totalRemoteDays,
        fDRemRate,
        finalDAgg.totalInOfficeDays,
        fDOffRate,
        finalDAgg.totalMainOfficeDays,
        finalDAgg.totalDowntownDays,
        finalDAgg.totalBranchDays,
        finalDAgg.totalClientSiteDays
      ]);
      rowTypeTracker.push("dept_total");
      chipCopyMapping.push(null);
    }

    var finalMAgg = dataModel.monthlyTotals[currentMonth];
    if (finalMAgg) {
      var fMCapUtil = (finalMAgg.totalContractCapacity > 0) ? (finalMAgg.totalWorkedDays / finalMAgg.totalContractCapacity) : 0;
      var fMRemRate = (finalMAgg.totalWorkedDays > 0) ? (finalMAgg.totalRemoteDays / finalMAgg.totalWorkedDays) : 0;
      var fMOffRate = (finalMAgg.totalWorkedDays > 0) ? (finalMAgg.totalInOfficeDays / finalMAgg.totalWorkedDays) : 0;

      tableValues.push([
        currentMonth,
        "All Departments",
        "Total Month (" + currentMonth + ")",
        finalMAgg.totalFte,
        finalMAgg.totalContractCapacity,
        finalMAgg.totalWorkedDays,
        finalMAgg.totalLeaveDays,
        fMCapUtil,
        finalMAgg.totalRemoteDays,
        fMRemRate,
        finalMAgg.totalInOfficeDays,
        fMOffRate,
        finalMAgg.totalMainOfficeDays,
        finalMAgg.totalDowntownDays,
        finalMAgg.totalBranchDays,
        finalMAgg.totalClientSiteDays
      ]);
      rowTypeTracker.push("month_total");
      chipCopyMapping.push(null);
    }
  }

  var numDataRows = tableValues.length;
  var totalRows = Math.max(numDataRows + 1, 20);

  // Clear existing sheet content
  kpiSheet.getRange(1, 1, Math.max(kpiSheet.getMaxRows(), totalRows), Math.max(kpiSheet.getMaxColumns(), numCols)).clearContent();

  // Write header row
  kpiSheet.getRange(1, 1, 1, numCols).setValues([headers]);
  var headerRange = kpiSheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#E8EAED");
  headerRange.setFontColor("#202124");
  headerRange.setVerticalAlignment("middle");
  kpiSheet.setFrozenRows(1);

  if (numDataRows > 0) {
    kpiSheet.getRange(2, 1, numDataRows, numCols).setValues(tableValues);

    // Copy People Chips for member rows
    if (membersSheet && typeof copyCellChip === "function") {
      for (var cr = 0; cr < chipCopyMapping.length; cr++) {
        var map = chipCopyMapping[cr];
        if (map && map.sourceRow > 0) {
          var targetRow = cr + 2;
          copyCellChip(membersSheet, map.sourceRow, 1, kpiSheet, targetRow, 3, map.name, ss.getId());
        }
      }
    }

    // Apply row styling for subtotal/total rows
    for (var rIdx = 0; rIdx < rowTypeTracker.length; rIdx++) {
      var rType = rowTypeTracker[rIdx];
      var sheetR = rIdx + 2;
      if (rType === "dept_total") {
        var deptRowRange = kpiSheet.getRange(sheetR, 1, 1, numCols);
        deptRowRange.setFontWeight("bold");
        deptRowRange.setBackground("#F1F3F4");
      } else if (rType === "month_total") {
        var monthRowRange = kpiSheet.getRange(sheetR, 1, 1, numCols);
        monthRowRange.setFontWeight("bold");
        monthRowRange.setBackground("#E2E7EC");
      }
    }

    // Number Formatting
    try {
      // Column 4 (FTE Rate): 0.00
      kpiSheet.getRange(2, 4, numDataRows, 1).setNumberFormat("0.00");
      // Column 5, 6, 7 (Capacity, Worked, Leave Days): 0.0
      kpiSheet.getRange(2, 5, numDataRows, 3).setNumberFormat("0.0");
      // Column 8 (Capacity Utilization): 0.0%
      kpiSheet.getRange(2, 8, numDataRows, 1).setNumberFormat("0.0%");
      // Column 9 (Remote Days): 0.0
      kpiSheet.getRange(2, 9, numDataRows, 1).setNumberFormat("0.0");
      // Column 10 (Remote Rate): 0.0%
      kpiSheet.getRange(2, 10, numDataRows, 1).setNumberFormat("0.0%");
      // Column 11 (In-Office Days): 0.0
      kpiSheet.getRange(2, 11, numDataRows, 1).setNumberFormat("0.0");
      // Column 12 (In-Office Rate): 0.0%
      kpiSheet.getRange(2, 12, numDataRows, 1).setNumberFormat("0.0%");
      // Columns 13-16 (Specific Offices): 0.0
      kpiSheet.getRange(2, 13, numDataRows, 4).setNumberFormat("0.0");
    } catch (eFmt) {}

    // Alignments
    try {
      kpiSheet.getRange(2, 1, numDataRows, 1).setHorizontalAlignment("center"); // Month
      kpiSheet.getRange(2, 2, numDataRows, 1).setHorizontalAlignment("left");   // Department
      kpiSheet.getRange(2, 3, numDataRows, 1).setHorizontalAlignment("left");   // Member
      kpiSheet.getRange(2, 4, numDataRows, numCols - 3).setHorizontalAlignment("right"); // Numbers & Percentages
    } catch (eAlign) {}
  }

  // Apply typography (12px Google Sans)
  if (typeof applyTypography === "function") {
    applyTypography(kpiSheet.getRange(1, 1, Math.max(numDataRows + 1, 2), numCols));
  } else {
    kpiSheet.getRange(1, 1, Math.max(numDataRows + 1, 2), numCols).setFontFamily("Google Sans").setFontSize(12);
  }

  // Column Widths
  try {
    kpiSheet.setColumnWidth(1, 110); // Month
    kpiSheet.setColumnWidth(2, 160); // Department
    kpiSheet.setColumnWidth(3, 190); // Member
    kpiSheet.setColumnWidth(4, 100); // FTE Rate
    kpiSheet.setColumnWidth(5, 170); // Contract Capacity (Days)
    kpiSheet.setColumnWidth(6, 120); // Worked Days
    kpiSheet.setColumnWidth(7, 110); // Leave Days
    kpiSheet.setColumnWidth(8, 160); // Capacity Utilization
    kpiSheet.setColumnWidth(9, 120); // Remote Days
    kpiSheet.setColumnWidth(10, 120); // Remote Rate
    kpiSheet.setColumnWidth(11, 120); // In-Office Days
    kpiSheet.setColumnWidth(12, 120); // In-Office Rate
    kpiSheet.setColumnWidth(13, 140); // Main Office (Days)
    kpiSheet.setColumnWidth(14, 150); // Downtown Office (Days)
    kpiSheet.setColumnWidth(15, 140); // Branch Office (Days)
    kpiSheet.setColumnWidth(16, 130); // Client Site (Days)
  } catch (eWidth) {}

  return {
    success: true,
    sheetName: "calc_kpis",
    memberRowsCount: memberRows.length,
    totalTableRows: numDataRows,
    monthsCount: dataModel.monthsCount
  };
}

/**
 * Public execution entry point to refresh calc_kpis dashboard.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function refreshKpis(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var res = renderKpiDashboard(ss);
  if (res.success && ss && typeof ss.toast === "function") {
    try {
      ss.toast(
        "Refreshed Presence & Remote KPIs: " + res.memberRowsCount + " member metrics across " + res.monthsCount + " months.",
        "KPIs Updated",
        5
      );
    } catch (eToast) {}
  }
  return res;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseKpiDate: parseKpiDate,
    formatMonthKey: formatMonthKey,
    countMonthWeekdays: countMonthWeekdays,
    isLeaveLocation: isLeaveLocation,
    categorizeLocation: categorizeLocation,
    computePresenceKpis: computePresenceKpis,
    renderKpiDashboard: renderKpiDashboard,
    refreshKpis: refreshKpis
  };
}
