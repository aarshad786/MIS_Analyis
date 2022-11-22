const { parse } = require("node-xlsx");
const { writeFileSync } = require("fs");
const moment = require("moment");
var markdownpdf = require("markdown-pdf");

const createMakrdown = (from, to) =>
  new Promise((resolve) => {
    markdownpdf().from(from).to(to, resolve);
  });

const worksheet = parse(`./classeslagging.xlsx`, { cellDates: true });

const convertSheetToObjects = (data) => {
  const objects = [];
  const keys = data[0];
  data.slice(1).forEach((x) => {
    let object = {};
    keys.forEach((key, i) => {
      object[key] = x[i];
    });
    objects.push(object);
  });
  return objects;
};

const sheet1Data = convertSheetToObjects(worksheet[0].data);

let classInformation = {};
sheet1Data.forEach((x) => {
  const key = [x.region, x.local, x.centre, x.standard, x.division].join("-");
  const classInfo = classInformation[key];
  if (!classInfo) {
    classInformation[key] = {
      ...x,
      offset: Math.floor(
        parseInt(x.standard_hour_daywise_milestones__expected_class_hours, 10) - parseInt(x.class_hrs, 10)
      ),
      lastClassBefore_15days: moment().diff(moment(x.last_session_date, "YYYY-MM-DD"), "day") > 15,
    };
  }
});

const classeNotFilledUp = Object.values(classInformation).filter((x) => x.lastClassBefore_15days);

const classesLagginBehind = Object.values(classInformation).filter((x) => x.offset > 0);

const Analysis = {
  classesNotBeingFilledUp: {
    secunderabad: classeNotFilledUp
      .filter((x) => x.local === "Secunderabad")
      .map((x) => ({
        classDetails: [x.local, x.centre, x.standard, x.division].join(" "),
        lastFilledDate: moment(x.last_session_date, "YYYY-MM-DD").isValid()
          ? moment(x.last_session_date, "YYYY-MM-DD").format("DD MMMM YYYY")
          : "**NOT STARTED**",
      })),
    hyderabad: classeNotFilledUp
      .filter((x) => x.local === "Hyderabad")
      .map((x) => ({
        classDetails: [x.local, x.centre, x.standard, x.division].join(" "),
        lastFilledDate: moment(x.last_session_date, "YYYY-MM-DD").isValid()
          ? moment(x.last_session_date, "YYYY-MM-DD").format("DD MMMM YYYY")
          : "**NOT STARTED**",
      })),
    bengaluru: classeNotFilledUp
      .filter((x) => x.local === "Bengaluru")
      .map((x) => ({
        classDetails: [x.local, x.centre, x.standard, x.division].join(" "),
        lastFilledDate: moment(x.last_session_date, "YYYY-MM-DD").isValid()
          ? moment(x.last_session_date, "YYYY-MM-DD").format("DD MMMM YYYY")
          : "**NOT STARTED**",
      })),
  },
  classesLaggingBehind: {
    secunderabad: classesLagginBehind
      .filter((x) => x.local === "Secunderabad")
      .map((x) => ({
        classDetails: [x.local, x.centre, x.standard, x.division].join(" "),
        laggingBy: x.offset,
      })),
    hyderabad: classesLagginBehind
      .filter((x) => x.local === "Hyderabad")
      .map((x) => ({
        classDetails: [x.local, x.centre, x.standard, x.division].join(" "),
        laggingBy: x.offset,
      })),
    bengaluru: classesLagginBehind
      .filter((x) => x.local === "Bengaluru")
      .map((x) => ({
        classDetails: [x.local, x.centre, x.standard, x.division].join(" "),
        laggingBy: x.offset,
      })),
  },
};

const GetReportMarkdown = (classNotFilled, classLagging) => [
  `**Summary**\n- Classes not Started **_${
    classNotFilled.filter((x) => x.lastFilledDate === "**NOT STARTED**").length
  }_**\n- Classes lagging Behind **_${classLagging.length}_**\n- Classes Not Regularly filled up **_${
    classNotFilled.length
  }_**`,
  "### Classes Lagging Behind",
  `| S.No. | Class | Lagging By _(hours)_ |\n| :---: | :---: | :---: |\n${classLagging
    .map((x, i) => "| " + (i + 1) + " | " + x.classDetails + " | " + x.laggingBy + " |")
    .join("\n")}`,
  "### Classes Not Filled Up in last 15 days",
  `| S.No. | Class | Last Filled Date |\n| :---: | :---: | :---: |\n${classNotFilled
    .map((x, i) => "| " + (i + 1) + " | " + x.classDetails + " | " + x.lastFilledDate + " |")
    .join("\n")}`,
];

writeFileSync(
  "./report.md",
  [
    `# Southern India Report - MIS (**_${moment().format("Do MMM YYYY")}_**)`,
    "## Hyderabad",
    ...GetReportMarkdown(Analysis.classesNotBeingFilledUp.hyderabad, Analysis.classesLaggingBehind.hyderabad),

    "## Secunderabad",
    ...GetReportMarkdown(Analysis.classesNotBeingFilledUp.secunderabad, Analysis.classesLaggingBehind.secunderabad),

    "## Bengaluru",
    ...GetReportMarkdown(Analysis.classesNotBeingFilledUp.bengaluru, Analysis.classesLaggingBehind.bengaluru),
  ].join("\n\n")
);

writeFileSync(
  "./Hyderabad-report.md",
  [
    `# Hyderabad (**_${moment().format("Do MMM YYYY")}_**)`,
    ...GetReportMarkdown(Analysis.classesNotBeingFilledUp.hyderabad, Analysis.classesLaggingBehind.hyderabad),
  ].join("\n\n")
);

writeFileSync(
  "./Secunderabad-report.md",
  [
    `# Secunderabad (**_${moment().format("Do MMM YYYY")}_**)`,
    ...GetReportMarkdown(Analysis.classesNotBeingFilledUp.secunderabad, Analysis.classesLaggingBehind.secunderabad),
  ].join("\n\n")
);

writeFileSync(
  "./Bengaluru-report.md",
  [
    `# Bengaluru (**_${moment().format("Do MMM YYYY")}_**)`,
    ...GetReportMarkdown(Analysis.classesNotBeingFilledUp.bengaluru, Analysis.classesLaggingBehind.bengaluru),
  ].join("\n\n")
);

writeFileSync(
  "./analysis.json",
  JSON.stringify(
    {
      analysis: Analysis,
      actuals: classInformation,
    },
    null,
    2
  )
);
