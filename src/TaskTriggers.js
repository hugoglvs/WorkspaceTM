/**
 * Task triggers (ADR 0051): events the app already notices, which create tasks
 * according to the Task Templates tab ("Created when", "Days after", "Assign to",
 * "Instructions").
 *
 * Adding an event = one entry in TASK_TRIGGERS (a label shown in the Task Types
 * dropdown) and one call to fireTaskTrigger(ss, key, context) where the app
 * notices the event. Everything else (which tasks, for whom, when) is set in
 * the Task Templates tab.
 *
 * Each task is created once per event: its Origin is
 * "trigger:<event>:<subject>:<type>", checked before creating.
 */

var TASK_TRIGGERS = {
  newMember: {
    label: "New member",
    help: "A person's first contract is declared. Date: the contract start. Person concerned: the new member."
  }
};

/** A function by name: the Apps Script global, else (Node tests) from its module. */
function trigFn(name, modulePath) {
  var g = (typeof globalThis !== "undefined") ? globalThis : this;
  if (g && typeof g[name] === "function") return g[name];
  if (typeof require !== "undefined") return require(modulePath)[name];
  return null;
}

/** Labels for the Task Types "Created when" dropdown. */
function taskTriggerLabels() {
  return Object.keys(TASK_TRIGGERS).map(function(k) { return TASK_TRIGGERS[k].label; });
}

/** Trigger key of a "Created when" label, or "" when none. */
function taskTriggerKey(label) {
  var l = String(label || "").trim().toLowerCase();
  if (!l) return "";
  for (var k in TASK_TRIGGERS) {
    if (TASK_TRIGGERS.hasOwnProperty(k) && (TASK_TRIGGERS[k].label.toLowerCase() === l || k.toLowerCase() === l)) return k;
  }
  return "";
}

/**
 * People a triggered task goes to: "Person concerned" (or blank) = the
 * event's person; otherwise names and "Role: X", as in Recurring Tasks.
 */
function triggerAssignees(assignTo, ctx, members) {
  var text = String(assignTo || "").trim();
  var out = [];
  var seen = {};
  var add = function(m) { if (m && !seen[m.key || m.name]) { seen[m.key || m.name] = true; out.push(m); } };
  var rest = [];
  (text ? text.split(/[,;\n]/) : ["person concerned"]).forEach(function(tok) {
    var t = tok.trim();
    if (!t) return;
    if (/^person( concerned)?$/i.test(t)) add(ctx.person);
    else rest.push(t);
  });
  if (rest.length) trigFn("resolveAppliesTo", "./RecurringTasks")(rest.join("; "), members).forEach(add);
  return out;
}

/**
 * Creates the tasks configured for an event, once each.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} eventKey a TASK_TRIGGERS key
 * @param {{subject: string, date: Date, person: (Object|undefined), project: (string|undefined)}} ctx
 *   subject identifies the occurrence (e.g. the person's name), date anchors "Days after"
 * @return {{created: number, skipped: number}}
 */
function fireTaskTrigger(ss, eventKey, ctx) {
  var res = { created: 0, skipped: 0 };
  if (!TASK_TRIGGERS[eventKey] || !ctx || !ctx.subject) return res;
  var types = trigFn("readTaskTypes", "./TaskTypes")(ss).filter(function(t) { return t.active && taskTriggerKey(t.createdWhen) === eventKey; });
  if (!types.length) return res;
  var existing = {};
  trigFn("readTaskRows", "./RecurringTasks")(ss).forEach(function(t) { if (t.origin) existing[t.origin.toLowerCase()] = true; });
  var members = trigFn("rtMembers", "./RecurringTasks")(ss);
  var createTask = trigFn("createTaskRow", "./RecurringTasks");
  var anchor = toDate(ctx.date) || toDate(new Date());
  types.forEach(function(type) {
    var origin = "trigger:" + eventKey + ":" + ctx.subject + ":" + type.name;
    if (existing[origin.toLowerCase()]) { res.skipped++; return; }
    var people = triggerAssignees(type.assignTo, ctx, members);
    if (!people.length) people = [null];
    people.forEach(function(person) {
      createTask(ss, {
        type: type.name, title: type.name, description: type.instructions || "",
        member: person, project: ctx.project || "",
        due: addDays(anchor, type.daysAfter || 0), origin: origin, assignedBy: "Automatic"
      });
      res.created++;
    });
  });
  return res;
}

/**
 * New member (ADR 0051): a person whose first contract starts today or later,
 * or started less than 30 days ago (so a run that missed it catches up). Fired
 * by the members sync and when the Contracts sidebar saves a first contract.
 *
 * @return {number} tasks created
 */
function fireNewMemberTriggers(ss, asOf) {
  var today = toDate(asOf || new Date());
  var since = addDays(today, -30);
  var first = {};
  trigFn("ctContracts", "./Contracts")(ss).forEach(function(c) {
    var start = toDate(c.startDate);
    if (start && (!first[c.key] || start < first[c.key])) first[c.key] = start;
  });
  var created = 0;
  trigFn("rtMembers", "./RecurringTasks")(ss).forEach(function(m) {
    var start = first[m.key];
    if (!start || start < since) return;
    created += fireTaskTrigger(ss, "newMember", { subject: m.name, date: start, person: m }).created;
    // The newcomer also takes their place in the rotations that include them (ADR 0060)
    try { trigFn("replanRotationsForNewcomer", "./RecurringTasks")(ss, m.key, today); } catch (eRot) { Logger.log("Rotation re-plan note: " + eRot); }
  });
  return created;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    TASK_TRIGGERS: TASK_TRIGGERS,
    taskTriggerLabels: taskTriggerLabels,
    taskTriggerKey: taskTriggerKey,
    triggerAssignees: triggerAssignees,
    fireTaskTrigger: fireTaskTrigger,
    fireNewMemberTriggers: fireNewMemberTriggers
  };
}
