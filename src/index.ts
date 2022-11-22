import { parse } from "node-xlsx";

const worksheet = parse(`./classeslagging.xlsx`, { cellDates: true });

const convertSheetToObjects = (data: unknown[][]) => {
  const objects: unknown[] = [];
  const keys = data[0] as string[];
  data.slice(1).forEach((x: unknown[]) => {
    let object: any = {};
    keys.forEach((key, i) => {
      object[key] = x[i];
    });
    objects.push(object);
  });
  return objects;
};

const sheet1Data = convertSheetToObjects(worksheet[0].data as unknown[][]);

let classInformation: any = {};
sheet1Data.forEach((x: any) => {
  const key = [x.region, x.local, x.centre, x.standard, x.division].join("-");
  const classInfo = classInformation[key];
  if (!classInfo) {
    classInformation[key] = {
      ...x,
      offset: parseInt(x.standard_hour_daywise_milestones__expected_class_hours, 10) - parseInt(x.class_hrs, 10),
    };
  }
});

interface IRowData {
  academic_year: string;
  mis_name: string;
  class_id: number;
  is_host: boolean;
  region: string;
  local: string;
  centre: string;
  standard: string;
  division: string;
  class_hrs: number;
  no_of_students: number;
  teacher_hr_id: number;
  teacher_first_name: string;
  teacher_last_name: string;
  last_session_date: string;
  standard_hour_daywise_milestones__id: number;
  standard_hour_daywise_milestones__cutoff_date: string;
  standard_hour_daywise_milestones__z3: number;
  standard_hour_daywise_milestones__z4: number;
  standard_hour_daywise_milestones__expected_class_hours: number;
}
