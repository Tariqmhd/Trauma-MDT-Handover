

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { motion, AnimatePresence } from "framer-motion";
import {
  Activity, AlertTriangle, BedDouble, CheckCircle2, ChevronRight,
  ClipboardCheck, Clock3, HeartPulse, Hospital, ListChecks, MapPin,
  Menu, MonitorUp, Search, ShieldCheck, Stethoscope, Syringe, Printer,
  Theater, UserRound, UsersRound, X, Plus, Pencil, FileSpreadsheet, Download, CheckSquare, Square, Save, Database
} from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "./components/ui/card";
import { Button } from "./components/ui/button";
import { Badge } from "./components/ui/badge";
import { Input } from "./components/ui/input";

const patients = [
  {id:"UAT-001",name:"SMITH, John",nhs:"9990000001",hospital:"H100001",dob:"12/03/1942",age:84,ward:"Trauma A",bed:"12",pathway:"NOF",diagnosis:"Left neck of femur fracture",priority:"Red",status:"Awaiting theatre",theatre:"Theatre 1",procedure:"Hemiarthroplasty",fourAT:8,amts:5,og:false,bone:false,falls:false,nhfd:"Incomplete",owner:"Registrar",due:"09:00",action:"Anaesthetic review; repeat potassium; confirm HDU availability",plan:"Optimise, complete anaesthetic assessment and confirm theatre order."},
  {id:"UAT-002",name:"JONES, Mary",nhs:"9990000002",hospital:"H100002",dob:"15/09/1938",age:88,ward:"Hazel",bed:"5",pathway:"NOF",diagnosis:"Intertrochanteric hip fracture",priority:"Amber",status:"Post-operative",theatre:"Completed",procedure:"DHS",fourAT:2,amts:8,og:true,bone:false,falls:true,nhfd:"Incomplete",owner:"NOF team",due:"11:00",action:"Bone health plan; discharge planning",plan:"Mobilise with therapy and complete secondary prevention."},
  {id:"UAT-003",name:"BROWN, Peter",nhs:"9990000003",hospital:"H100003",dob:"17/04/1962",age:64,ward:"ED Majors",bed:"M4",pathway:"General trauma",diagnosis:"Suspected cauda equina syndrome",priority:"Red",status:"Awaiting MRI",theatre:"Not listed",procedure:"Pending imaging",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"SHO",due:"09:15",action:"Confirm MRI slot; repeat neurology; contact spinal service",plan:"Urgent MRI and spinal review."},
  {id:"UAT-004",name:"TAYLOR, Emma",nhs:"9990000004",hospital:"H100004",dob:"21/08/1995",age:31,ward:"Trauma B",bed:"7",pathway:"General trauma",diagnosis:"Closed bimalleolar ankle fracture",priority:"Amber",status:"Listed",theatre:"Theatre 2",procedure:"Ankle ORIF",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Operating team",due:"10:30",action:"Confirm implants; mark patient",plan:"Proceed on trauma list when ready."},
  {id:"UAT-005",name:"WILSON, David",nhs:"9990000005",hospital:"H100005",dob:"05/02/1954",age:72,ward:"Trauma A",bed:"14",pathway:"General trauma",diagnosis:"Tibial plateau fracture",priority:"Amber",status:"Awaiting theatre",theatre:"Backup",procedure:"ORIF",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Registrar",due:"12:00",action:"Check CT plan; consent",plan:"Finalise fixation strategy and theatre date."},
  {id:"UAT-006",name:"THOMAS, Sarah",nhs:"9990000006",hospital:"H100006",dob:"09/10/1940",age:85,ward:"Hazel",bed:"2",pathway:"NOF",diagnosis:"Displaced intracapsular hip fracture",priority:"Red",status:"Theatre today",theatre:"Theatre 1",procedure:"Hemiarthroplasty",fourAT:6,amts:6,og:false,bone:false,falls:false,nhfd:"Incomplete",owner:"Registrar",due:"09:30",action:"Orthogeriatric review; anaesthetic optimisation",plan:"Proceed following optimisation and senior review."},
  {id:"UAT-007",name:"ROBERTS, Anne",nhs:"9990000007",hospital:"H100007",dob:"18/06/1945",age:81,ward:"Trauma A",bed:"10",pathway:"NOF",diagnosis:"Extracapsular hip fracture",priority:"Green",status:"Post-operative",theatre:"Completed",procedure:"DHS",fourAT:0,amts:10,og:true,bone:true,falls:true,nhfd:"Ready",owner:"Ward team",due:"14:00",action:"Therapy review; confirm discharge destination",plan:"Mobilise and progress discharge pathway."},
  {id:"UAT-008",name:"EVANS, Michael",nhs:"9990000008",hospital:"H100008",dob:"22/07/1978",age:48,ward:"Trauma B",bed:"6",pathway:"General trauma",diagnosis:"Distal radius fracture",priority:"Green",status:"Post-operative",theatre:"Completed",procedure:"Distal radius ORIF",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Ward team",due:"15:00",action:"Post-op imaging; discharge medication",plan:"Post-operative checks and discharge."},
  {id:"UAT-009",name:"HUGHES, Linda",nhs:"9990000009",hospital:"H100009",dob:"14/05/1956",age:70,ward:"Trauma A",bed:"21",pathway:"General trauma",diagnosis:"Periprosthetic femoral fracture",priority:"Red",status:"Awaiting theatre",theatre:"Theatre 2",procedure:"Revision fixation",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Trauma coordinator",due:"10:00",action:"Blood crossmatch; implant check; consultant plan",plan:"Confirm implants, blood availability and theatre order."},
  {id:"UAT-010",name:"DAVIES, Brian",nhs:"9990000010",hospital:"H100010",dob:"01/01/1943",age:83,ward:"Hazel",bed:"8",pathway:"NOF",diagnosis:"Subtrochanteric femoral fracture",priority:"Amber",status:"Awaiting theatre",theatre:"Backup",procedure:"Long nail",fourAT:4,amts:6,og:false,bone:false,falls:true,nhfd:"Incomplete",owner:"Registrar",due:"12:00",action:"Consent; orthogeriatric review",plan:"Complete readiness checks and retain as backup case."},
  {id:"UAT-011",name:"HARRIS, Susan",nhs:"9990000011",hospital:"H100011",dob:"20/02/1952",age:74,ward:"Trauma B",bed:"9",pathway:"General trauma",diagnosis:"Open tibial fracture",priority:"Red",status:"Theatre planned",theatre:"Theatre 1",procedure:"Debridement and fixation",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Registrar",due:"09:45",action:"Antibiotics; neurovascular checks; theatre readiness",plan:"Urgent debridement and stabilisation."},
  {id:"UAT-012",name:"LEWIS, James",nhs:"9990000012",hospital:"H100012",dob:"18/12/1982",age:43,ward:"Trauma A",bed:"19",pathway:"General trauma",diagnosis:"Olecranon fracture",priority:"Green",status:"Awaiting clinic",theatre:"Not listed",procedure:"Pending decision",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"SHO",due:"13:00",action:"Confirm upper limb plan; clinic slot",plan:"Discuss at upper-limb review."},
  {id:"UAT-013",name:"CLARKE, Patricia",nhs:"9990000013",hospital:"H100013",dob:"24/08/1937",age:89,ward:"Hazel",bed:"11",pathway:"NOF",diagnosis:"Intracapsular hip fracture",priority:"Amber",status:"Post-operative",theatre:"Completed",procedure:"Hemiarthroplasty",fourAT:1,amts:8,og:true,bone:true,falls:true,nhfd:"Ready",owner:"NOF team",due:"11:30",action:"Mobilise; confirm NHFD review data",plan:"Mobilise and verify pathway completion."},
  {id:"UAT-014",name:"WALKER, Daniel",nhs:"9990000014",hospital:"H100014",dob:"13/09/1971",age:55,ward:"ED Majors",bed:"M2",pathway:"General trauma",diagnosis:"Suspected septic arthritis",priority:"Red",status:"Awaiting review",theatre:"Possible",procedure:"Joint washout",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Registrar",due:"09:00",action:"Aspiration; blood cultures; consultant review",plan:"Urgent senior assessment and infection work-up."},
  {id:"UAT-015",name:"HALL, Barbara",nhs:"9990000015",hospital:"H100015",dob:"22/11/1935",age:90,ward:"Trauma A",bed:"17",pathway:"NOF",diagnosis:"Intertrochanteric fracture",priority:"Amber",status:"Awaiting theatre",theatre:"Backup",procedure:"DHS",fourAT:5,amts:4,og:false,bone:false,falls:false,nhfd:"Incomplete",owner:"Registrar",due:"10:15",action:"Anaesthetic review; consent; delirium plan",plan:"Optimise and prepare as backup patient."},
  {id:"UAT-016",name:"ALLEN, Richard",nhs:"9990000016",hospital:"H100016",dob:"10/06/1967",age:59,ward:"Trauma B",bed:"15",pathway:"General trauma",diagnosis:"Calcaneal fracture",priority:"Green",status:"Awaiting surgery",theatre:"Not listed",procedure:"ORIF",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Foot and ankle team",due:"14:30",action:"Soft tissue review; CT discussion",plan:"Await soft-tissue recovery and definitive plan."},
  {id:"UAT-017",name:"YOUNG, Christine",nhs:"9990000017",hospital:"H100017",dob:"19/04/1948",age:78,ward:"Hazel",bed:"6",pathway:"NOF",diagnosis:"Intertrochanteric fracture",priority:"Amber",status:"Post-operative",theatre:"Completed",procedure:"DHS",fourAT:2,amts:9,og:true,bone:false,falls:true,nhfd:"Incomplete",owner:"NOF team",due:"12:30",action:"Bone health assessment; therapy review",plan:"Progress therapy and secondary prevention."},
  {id:"UAT-018",name:"KING, Anthony",nhs:"9990000018",hospital:"H100018",dob:"30/01/1950",age:76,ward:"Trauma A",bed:"1",pathway:"General trauma",diagnosis:"Acetabular fracture",priority:"Red",status:"Awaiting transfer",theatre:"Tertiary centre",procedure:"Specialist fixation",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Consultant",due:"09:30",action:"Confirm tertiary centre acceptance; transfer imaging",plan:"Coordinate specialist transfer."},
  {id:"UAT-019",name:"SCOTT, Margaret",nhs:"9990000019",hospital:"H100019",dob:"08/07/1936",age:90,ward:"Hazel",bed:"4",pathway:"NOF",diagnosis:"Displaced intracapsular hip fracture",priority:"Red",status:"Awaiting theatre",theatre:"Theatre 2",procedure:"Hemiarthroplasty",fourAT:7,amts:4,og:false,bone:false,falls:false,nhfd:"Incomplete",owner:"Consultant",due:"11:00",action:"Orthogeriatric review; anaesthetic optimisation; consent",plan:"Urgent optimisation and theatre planning."},
  {id:"UAT-020",name:"GREEN, Christopher",nhs:"9990000020",hospital:"H100020",dob:"15/03/1988",age:38,ward:"Trauma B",bed:"2",pathway:"General trauma",diagnosis:"Bimalleolar ankle fracture",priority:"Amber",status:"Listed",theatre:"Theatre 2",procedure:"Ankle ORIF",fourAT:null,amts:null,og:null,bone:null,falls:null,nhfd:"N/A",owner:"Operating team",due:"10:45",action:"Confirm list order; mark patient",plan:"Proceed when theatre ready."},
  {id:"UAT-021",name:"MORGAN, Helen",nhs:"9990000021",hospital:"H100021",dob:"11/10/1941",age:84,ward:"Hazel",bed:"12",pathway:"NOF",diagnosis:"Intertrochanteric hip fracture",priority:"Green",status:"Post-operative",theatre:"Completed",procedure:"DHS",fourAT:1,amts:9,og:true,bone:true,falls:true,nhfd:"Ready",owner:"NOF team",due:"15:00",action:"Continue mobilisation and discharge planning",plan:"Progress rehabilitation and complete NHFD follow-up data."}
];

const nav = [
  ["Command Centre", MonitorUp], ["Patients Admission List", ListChecks], ["Patient Detail", UserRound],
  ["NOF Dashboard", HeartPulse], ["Theatre Board", Theater], ["MDT / Handover", UsersRound]
];

// Ten fictional NOF records with populated NHFD v16 fields for dashboard/UAT demonstration.
const sampleNHFD=(p,o={})=>{
  const presentationDate=o.presentationDate, presentationTime=o.presentationTime||"09:00";
  const wardAdmissionDate=o.wardAdmissionDate||presentationDate, wardAdmissionTime=o.wardAdmissionTime||"12:00";
  const surgeryDate=o.surgeryDate||presentationDate, surgeryTime=o.surgeryTime||"20:00";
  const geriatricianDate=o.geriatricianDate||presentationDate, geriatricianTime=o.geriatricianTime||"15:00";
  const postGood=o.postGood!==false, pass=o.pass!==false;
  return {
    operatingHospitalCode:"RWP",nhsNumber:p.nhs,hospitalNumber:p.hospital,firstName:p.name.split(", ")[1]||"",surname:p.name.split(", ")[0]||"",dob:p.dob,sex:o.sex||"2",postcode:o.postcode||"WR1 1AA",firstHospitalCode:"X",
    presentationDate,presentationTime,ambulance:"1",ambulanceDate:presentationDate,ambulanceTime:o.ambulanceTime||"08:30",residenceBeforeAdmission:"1",presentationViaED:"1",admittedToOrthoWard:"1",wardAdmissionDate,wardAdmissionTime,nerveBlockED:"1",
    mobilityPreFracture:"3",preop4AT:"2",preop4ATAlertness:pass?"0":"1",preop4ATAMT4:pass?"0":"4",preop4ATAttention:pass?"0":"2",preop4ATAcuteChange:pass?"0":"4",asaGrade:"2",nutritionRisk:pass?"1":"0",boneProtectionPre:"2",
    fractureType:o.fractureType||"14",fractureSide:o.fractureSide||"1",pathologicalFracture:"0",operated:"1",surgeryDate,surgeryTime,operationType:o.operationType||"1",anaesthesia:["1","2"],delayReason:pass?"0":"3",scrubbedSurgeonGrade:"1",surgeonGrade:"1",anaesthetistGrade:"1",weightBearPostOp:"1",
    physioAssessment:pass?"1":"0",mobilisedPostSurgery:pass?"1":"0",geriatricianGrade:"1",geriatricianDate,geriatricianTime,fallsAssessment:pass?"1":"0",pressureUlcers:"0",medicationPostFracture:pass?"1":"0",
    postop4AT:"2",postop4ATAlertness:postGood?"0":"1",postop4ATAMT4:postGood?"0":"4",postop4ATAttention:postGood?"0":"2",postop4ATAcuteChange:postGood?"0":"4",
    wardDischargeDate:o.wardDischargeDate||presentationDate,wardDischargeDestination:"1",trustDischargeDate:o.trustDischargeDate||presentationDate,trustDischargeDestination:"1",reoperations:"0",residentialStatus:"1",postFractureMobility:"3",medicationAtFollowup:"1"
  };
};
const sampleNHFDById={
  "UAT-001":sampleNHFD(patients.find(x=>x.id==="UAT-001"),{presentationDate:"05/01/2026",wardAdmissionTime:"12:30",surgeryDate:"06/01/2026",surgeryTime:"18:00",geriatricianDate:"06/01/2026",geriatricianTime:"14:00",pass:true}),
  "UAT-002":sampleNHFD(patients.find(x=>x.id==="UAT-002"),{presentationDate:"03/02/2026",surgeryDate:"04/02/2026",surgeryTime:"16:00",pass:true}),
  "UAT-006":sampleNHFD(patients.find(x=>x.id==="UAT-006"),{presentationDate:"11/03/2026",surgeryDate:"13/03/2026",surgeryTime:"21:00",pass:false,postGood:false}),
  "UAT-007":sampleNHFD(patients.find(x=>x.id==="UAT-007"),{presentationDate:"08/04/2026",surgeryDate:"08/04/2026",surgeryTime:"17:00",pass:true}),
  "UAT-010":sampleNHFD(patients.find(x=>x.id==="UAT-010"),{presentationDate:"19/05/2026",surgeryDate:"20/05/2026",surgeryTime:"17:00",pass:false}),
  "UAT-013":sampleNHFD(patients.find(x=>x.id==="UAT-013"),{presentationDate:"14/06/2026",surgeryDate:"14/06/2026",surgeryTime:"15:00",pass:true}),
  "UAT-015":sampleNHFD(patients.find(x=>x.id==="UAT-015"),{presentationDate:"22/07/2026",surgeryDate:"24/07/2026",surgeryTime:"18:00",pass:false,postGood:false}),
  "UAT-017":sampleNHFD(patients.find(x=>x.id==="UAT-017"),{presentationDate:"16/08/2026",surgeryDate:"16/08/2026",surgeryTime:"14:00",pass:true}),
  "UAT-019":sampleNHFD(patients.find(x=>x.id==="UAT-019"),{presentationDate:"25/09/2026",surgeryDate:"26/09/2026",surgeryTime:"13:00",pass:true}),
  "UAT-021":sampleNHFD(patients.find(x=>x.id==="UAT-021"),{presentationDate:"01/10/2026",surgeryDate:"01/10/2026",surgeryTime:"16:00",pass:true})
};
patients.forEach(p=>{if(sampleNHFDById[p.id]) p.nhfdData=sampleNHFDById[p.id];});
const pClass={Red:"!bg-red-600 !text-white !border-red-600",Amber:"!bg-amber-500 !text-white !border-amber-500",Green:"!bg-emerald-600 !text-white !border-emerald-600"};
const pct=(n,d)=>d?Math.round(n/d*100):0;


const NHFD_V16_HEADERS = [
  "Operating hospital code","NHS number","Hospital number or patient ID","First name","Surname","Date Of Birth","Sex / Gender","Patients postcode",
  "First hospital code","Presentation date","Presentation time","Ambulance","Ambulance date","Ambulance time","Residence before admission",
  "Presentation via ED","Admitted to ortho ward","Date ward admission","Time ward admission","Nerve Block ED","Mobility pre fracture","Pre-op 4AT","Pre-op 4AT Alertness","Pre-op 4AT AMT4","Pre-op 4AT Attention","Pre-op 4AT Acute change",
  "ASA Grade","Nutrition risk assessment","Bone protection medication pre fracture","Fracture type","Fracture side","Pathological fracture","Operated","Surgery date","Surgery time","Operation type","Anaesthesia","Delay reason","Scrubbed surgeon grade","Surgeon grade","Anaesthetist grade","Weight-bear post-op",
  "Physiotherapist assessment","Mobilised post surgery","Geriatrician grade","Geriatrician date","Geriatrician time","Falls assessment","Pressure ulcers","Medication post fracture","Post-op 4AT","4AT Alertness","4AT AMT4","4AT Attention","4AT Acute change",
  "Ward discharge date","Ward discharge destination","Trust discharge date","Trust discharge destination","Re-operations","Residential status","Post fracture mobility","Medication at followup"
];

const NHFD_DEFAULTS={
  ward:"", bed:"", diagnosis:"", priority:"Amber", status:"Awaiting theatre", theatre:"Not listed", procedure:"Pending", owner:"NOF team", due:"", action:"", plan:"",
  operatingHospitalCode:"RWP", nhsNumber:"", hospitalNumber:"", firstName:"", surname:"", dob:"", sex:"", postcode:"",
  firstHospitalCode:"X", presentationDate:"", presentationTime:"", ambulance:"", ambulanceDate:"", ambulanceTime:"",
  residenceBeforeAdmission:"", presentationViaED:"", admittedToOrthoWard:"", wardAdmissionDate:"", wardAdmissionTime:"", nerveBlockED:"",
  mobilityPreFracture:"", preop4AT:"", preop4ATAlertness:"", preop4ATAMT4:"", preop4ATAttention:"", preop4ATAcuteChange:"",
  asaGrade:"", nutritionRisk:"", boneProtectionPre:"", fractureType:"", fractureSide:"", pathologicalFracture:"",
  operated:"", surgeryDate:"", surgeryTime:"", operationType:"", anaesthesia:[], delayReason:"", scrubbedSurgeonGrade:"",
  surgeonGrade:"", anaesthetistGrade:"", weightBearPostOp:"", physioAssessment:"", mobilisedPostSurgery:"", geriatricianGrade:"",
  geriatricianDate:"", geriatricianTime:"", fallsAssessment:"", pressureUlcers:"", medicationPostFracture:"", postop4AT:"",
  postop4ATAlertness:"", postop4ATAMT4:"", postop4ATAttention:"", postop4ATAcuteChange:"", wardDischargeDate:"",
  wardDischargeDestination:"", trustDischargeDate:"", trustDischargeDestination:"", reoperations:"", residentialStatus:"",
  postFractureMobility:"", medicationAtFollowup:""
};
const NHFD_OPTIONS={
 sex:[["2","Female"],["1","Male"]],
 ambulance:[["0","No"],["1","Yes"]],
 residenceBeforeAdmission:[["1","Own home/sheltered housing"],["2","Residential care"],["3","Nursing care"]],
 presentationViaED:[["1","Yes"],["2","No - already inpatient on this hospital site"],["3","No - already inpatient in another hospital site of this Trust/HB"],["4","No - already inpatient in another Trust/HB"]],
 admittedToOrthoWard:[["1","Yes - admitted to orthopaedic/orthogeriatric ward"],["0","No - admitted from ED but never reached orthopaedic/orthogeriatric ward"],["2","No - not admitted from ED"]],
 nerveBlockED:[["1","Yes - by ambulance staff"],["2","Yes - in Emergency Department"],["3","Yes - in ward before going to theatre"],["0","Not done/not documented"]],
 mobilityPreFracture:[["1","Freely mobile without aids"],["2","Mobile outdoors with one aid"],["3","Mobile outdoors with two aids or frame"],["4","Some indoor mobility but never goes outside without help"],["5","No functional mobility"],["U","Unknown"]],
 preop4AT:[["2","4AT assessment prior to operation"],["0","Not done / Patient refused"]],
 fourATAlertness:[["0","Normal"],["4","Clearly abnormal"]],
 fourATAMT4:[["0","No mistakes"],["1","One mistake"],["2","Two or more mistakes / untestable"]],
 fourATAttention:[["0","No mistakes"],["1","Any mistakes or refuses"],["2","Too unwell or drowsy to test"]],
 fourATAcute:[["0","No change"],["4","Change"]],
 asaGrade:[["1","1 - Normal healthy patient"],["2","2 - Mild systemic disease"],["3","3 - Severe systemic disease"],["4","4 - Severe systemic disease, constant threat to life"],["5","5 - Moribund"],["U","Unknown"]],
 nutritionRisk:[["0","No"],["1","Yes - assessment indicates normal"],["2","Yes - malnourished"],["3","Yes - at risk of malnutrition"]],
 boneProtectionPre:[["9","Abaloparatide"],["1","Alendronate"],["7","Alfacalcidol or Calcitriol"],["6","Denosumab"],["3","Ibandronate"],["2","Risedronate"],["8","Romosozumab"],["5","Teriparatide"],["4","Zoledronate"],["0","Not taking any of these bone treatments"]],
 fractureType:[["11","Intracapsular - displaced"],["12","Intracapsular - undisplaced"],["13","Trochanteric"],["14","Intertrochanteric / Reverse oblique"],["15","Subtrochanteric"],["21","Femoral shaft fracture"],["31","Distal femoral fracture"],["41","Peri-prosthetic fracture - around a hip prosthesis"],["44","Peri-prosthetic fracture - around a knee replacement"],["45","Peri-prosthetic fracture - mid shaft"],["51","Pelvic fracture"]],
 fractureSide:[["1","Left"],["2","Right"]], pathologicalFracture:[["0","No"],["1","Yes"]],
 operated:[["1","Yes"],["2","No - Surgery not indicated for this fracture"],["3","No - Surgery not possible for this patient"],["4","No - Patient died before surgery could take place"]],
 operationType:[["1","Internal fixation - Sliding Hip Screw"],["2","Internal fixation - Cannulated screws"],["3","Internal fixation - IM nail (long)"],["4","Internal fixation - IM nail (short)"],["6","Arthroplasty - Hemiarthroplasty (cemented)"],["7","Arthroplasty - Hemiarthroplasty (uncemented)"],["8","Arthroplasty - Primary THR (uncemented)"],["9","Arthroplasty - Primary THR (cemented)"],["11","Revision THR"],["12","Revision TKR"],["13","Internal fixation - plate"],["15","Internal fixation - IM nail and plate"],["16","Revision THR and internal fixation"],["17","Revision TKR and internal fixation"],["18","Proximal femoral replacement"],["19","Distal femoral replacement"],["20","Total femoral replacement"],["21","Surgery for pelvic fracture"],["99","Other"]],
 anaesthesia:[["1","General"],["2","Spinal"],["3","Epidural"],["4","Intra-operative sedation"],["5","Intra-operative nerve-block"],["6","High volume peri-articular LA infiltration"]],
 delayReason:[["0","No delay"],["1","Awaiting fracture diagnosis or confirmation"],["2","Awaiting medical review/investigation or stabilisation"],["3","Delayed due to insufficient theatre capacity"],["5","Delayed for reversal of warfarin"],["6","Delayed for reversal of DOAC"],["8","Delayed due to appropriate surgeon/equipment not available"],["9","Other"],["U","Unknown"]],
 clinicianGrade:[["1","Consultant"],["2","Specialist/associate specialist"],["3","Staff-grade/specialty doctor"],["4","ST3+"],["5","Below ST3"],["U","Unknown"]],
 anaesthetistGrade:[["1","Consultant"],["2","Specialist/associate specialist"],["3","Staff-grade/specialty doctor"],["4","ST3+"],["5","Below ST3"],["6","Anaesthetic Associate"],["U","Unknown"]],
 weightBearPostOp:[["1","Yes"],["0","No"]], physioAssessment:[["1","Yes"],["0","No"]],
 mobilisedPostSurgery:[["1","Yes - physiotherapist"],["2","Yes - other ward staff"],["3","No - inadequate post-op. pain control"],["4","No - symptomatic hypotension"],["5","No - patient too agitated or confused"],["6","No - other documented clinical contraindication"],["7","No - lack of staff or other resources"],["8","No - other"]],
 geriatricianGrade:[["0","Not seen"],["1","Consultant"],["2","Specialist/associate specialist"],["3","Staff-grade/specialty doctor"],["4","ST3+"],["5","Below ST3"],["U","Unknown"]],
 fallsAssessment:[["1","Yes"],["0","No"]], pressureUlcers:[["0","No"],["1","Yes"],["U","Unknown"]],
 medicationPostFracture:[["9","Abaloparatide"],["1","Alendronate"],["7","Alfacalcidol or Calcitriol"],["6","Denosumab"],["3","Ibandronate"],["2","Risedronate"],["8","Romosozumab"],["5","Teriparatide"],["4","Zoledronate"],["10","Assessed - no bone protection medication needed/appropriate"],["11","Informed decline - patient decided not to take offered treatment"],["12","On no treatment - pending DXA scan or bone clinic assessment"],["0","No assessment or action taken"]],
 postop4AT:[["1","Assessed before the 3rd day after surgery"],["2","Assessed before the 7th day after surgery"],["3","Not done by the 7th day after surgery"],["0","Not done / Patient refused"]],
 wardDischargeDestination:[["1","Own home/sheltered housing"],["2","Residential care"],["3","Nursing care"],["4","Rehabilitation unit - hospital bed in this Trust/HB"],["5","Rehabilitation unit - hospital bed in another Trust/HB"],["6","Rehabilitation unit - NHS funded care home bed"],["7","Acute hospital"],["8","Dead"],["8a","Died pre surgery"],["9","Other"]],
 trustDischargeDestination:[["1","Own home/sheltered housing"],["2","Residential care"],["3","Nursing care"],["5","Rehabilitation unit - hospital bed in another Trust/HB"],["6","Rehabilitation unit - NHS funded care home bed"],["7","Acute hospital"],["8","Dead"],["8a","Died pre surgery"],["8b","Dead at ward discharge"],["9","Other"],["U","Unknown"]],
 reoperations:[["1","Yes"],["0","No"]], residentialStatus:[["1","Own home/sheltered housing"],["2","Residential care"],["3","Nursing care"],["4","Rehabilitation unit - hospital bed in this Trust"],["5","Rehabilitation unit - hospital bed in another Trust"],["6","Rehabilitation unit - NHS funded care home bed"],["7","Acute hospital"],["8","Dead"],["9","Other"],["U","Unknown"]],
 postFractureMobility:[["1","Freely mobile without aids"],["2","Mobile outdoors with one aid"],["3","Mobile outdoors with two aids or frame"],["4","Some indoor mobility but never goes outside without help"],["5","No functional mobility"],["U","Unknown"]],
 medicationAtFollowup:[["1","Yes - continues same bone therapy"],["2","Yes - now on another bone therapy, started after discharge"],["3","No longer appropriate (stopped by clinician)"],["4","No longer taking therapy (stopped by patient)"],["5","No bone therapy started"]]
};
const calculateAgeFromDob=(dob)=>{
  const m=String(dob||"").match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if(!m) return "";
  const d=new Date(Number(m[3]),Number(m[2])-1,Number(m[1])); if(Number.isNaN(d.getTime())) return "";
  const now=new Date(); let age=now.getFullYear()-d.getFullYear(); const md=now.getMonth()-d.getMonth(); if(md<0 || (md===0 && now.getDate()<d.getDate())) age--; return age;
};
const calculate4ATScore=(d={})=>{
  if(d.preop4AT!=="2") return null;
  const vals=[d.preop4ATAlertness,d.preop4ATAMT4,d.preop4ATAttention,d.preop4ATAcuteChange];
  if(vals.some(v=>String(v??"")==="")) return null;
  return Number(vals[0]||0)+Number(vals[1]||0)+Number(vals[2]||0)+Number(vals[3]||0);
};

const nhfdFullRow=(d={})=>[
  d.operatingHospitalCode||"",d.nhsNumber||"",d.hospitalNumber||"",d.firstName||"",d.surname||"",d.dob||"",d.sex||"",d.postcode||"",d.firstHospitalCode||"",d.presentationDate||"",d.presentationTime||"",d.ambulance||"",d.ambulanceDate||"",d.ambulanceTime||"",d.residenceBeforeAdmission||"",d.presentationViaED||"",d.admittedToOrthoWard||"",d.wardAdmissionDate||"",d.wardAdmissionTime||"",d.nerveBlockED||"",d.mobilityPreFracture||"",d.preop4AT||"",d.preop4ATAlertness||"",d.preop4ATAMT4||"",d.preop4ATAttention||"",d.preop4ATAcuteChange||"",d.asaGrade||"",d.nutritionRisk||"",d.boneProtectionPre||"",d.fractureType||"",d.fractureSide||"",d.pathologicalFracture||"",d.operated||"",d.surgeryDate||"",d.surgeryTime||"",d.operationType||"",Array.isArray(d.anaesthesia)?d.anaesthesia.join(";"):d.anaesthesia||"",d.delayReason||"",d.scrubbedSurgeonGrade||"",d.surgeonGrade||"",d.anaesthetistGrade||"",d.weightBearPostOp||"",d.physioAssessment||"",d.mobilisedPostSurgery||"",d.geriatricianGrade||"",d.geriatricianDate||"",d.geriatricianTime||"",d.fallsAssessment||"",d.pressureUlcers||"",d.medicationPostFracture||"",d.postop4AT||"",d.postop4ATAlertness||"",d.postop4ATAMT4||"",d.postop4ATAttention||"",d.postop4ATAcuteChange||"",d.wardDischargeDate||"",d.wardDischargeDestination||"",d.trustDischargeDate||"",d.trustDischargeDestination||"",d.reoperations||"",d.residentialStatus||"",d.postFractureMobility||"",d.medicationAtFollowup||""
];

const splitPatientName=(name)=>{
  const [surname,...rest]=String(name||"").split(",");
  return {surname:(surname||"").trim(), firstName:rest.join(",").trim()};
};
const inferFractureType=(diagnosis)=>{
  const d=String(diagnosis||"").toLowerCase();
  if(d.includes("intertrochanteric")) return "14";
  if(d.includes("subtrochanteric")) return "15";
  if(d.includes("intracapsular") && d.includes("displaced")) return "11";
  if(d.includes("intracapsular")) return "12";
  if(d.includes("neck of femur")) return "";
  if(d.includes("periprosthetic") && d.includes("knee")) return "44";
  if(d.includes("periprosthetic") || d.includes("peri-prosthetic")) return "41";
  if(d.includes("pelvic") || d.includes("acetabular")) return "51";
  if(d.includes("distal femoral")) return "31";
  if(d.includes("femoral shaft")) return "21";
  return "";
};
const inferSide=(diagnosis)=>{
  const d=String(diagnosis||"").toLowerCase();
  if(d.includes("left")) return "L";
  if(d.includes("right")) return "R";
  return "";
};
const inferOperation=(procedure)=>{
  const d=String(procedure||"").toLowerCase();
  if(d.includes("dhs")) return "1";
  if(d.includes("long nail")) return "3";
  return "";
};
const nhfdRow=(p)=>{
  if(p.nhfdData) return nhfdFullRow(p.nhfdData);
  const {surname,firstName}=splitPatientName(p.name);
  const operated=["Post-operative","Theatre today","Theatre planned"].includes(p.status)||p.theatre==="Completed";
  return nhfdFullRow({
    operatingHospitalCode:"RWP", nhsNumber:p.nhs, hospitalNumber:p.hospital, firstName, surname, dob:p.dob,
    firstHospitalCode:"X", preop4AT:p.fourAT!==null&&p.fourAT!==undefined?"2":"0", fractureType:inferFractureType(p.diagnosis),
    fractureSide:inferSide(p.diagnosis)==="L"?"1":inferSide(p.diagnosis)==="R"?"2":"", operated:operated?"1":"",
    operationType:inferOperation(p.procedure), fallsAssessment:p.falls===true?"1":p.falls===false?"0":""
  });
};
const nhfdMissing=(row)=>{
  const essential=[0,1,5,6,9,10,29,30];
  return essential.filter(i=>!String(row[i]??"").trim()).map(i=>NHFD_V16_HEADERS[i]);
};
const nhfdValidation=(d)=>{
  const row=nhfdFullRow(d), missing=[];
  const essential=[0,1,5,6,9,10,29,30];
  essential.forEach(i=>{if(!String(row[i]??"").trim()) missing.push(NHFD_V16_HEADERS[i]);});
  const required=["hospitalNumber","firstName","surname","presentationViaED","admittedToOrthoWard","nerveBlockED","mobilityPreFracture","preop4AT","asaGrade","nutritionRisk","boneProtectionPre","pathologicalFracture","operated"];
  required.forEach(k=>{if(!String(d[k]??"").trim()) missing.push(k);});
  if(d.ambulance==="1"){if(!d.ambulanceDate) missing.push("Ambulance date"); if(!d.ambulanceTime) missing.push("Ambulance time");}
  if(d.admittedToOrthoWard==="1"){if(!d.wardAdmissionDate) missing.push("Date ward admission"); if(!d.wardAdmissionTime) missing.push("Time ward admission");}
  if(d.preop4AT==="2"){["preop4ATAlertness","preop4ATAMT4","preop4ATAttention","preop4ATAcuteChange"].forEach(k=>{if(!String(d[k]??"").trim()) missing.push(k);});}
  if(d.operated==="1"){
    ["surgeryDate","surgeryTime","operationType","anaesthesia","delayReason","scrubbedSurgeonGrade","surgeonGrade","anaesthetistGrade","weightBearPostOp","physioAssessment","mobilisedPostSurgery","geriatricianGrade","fallsAssessment","pressureUlcers","medicationPostFracture","postop4AT","reoperations"].forEach(k=>{const v=d[k];if((Array.isArray(v)&&!v.length)||(!Array.isArray(v)&&!String(v??"").trim())) missing.push(k);});
    if(d.geriatricianGrade && d.geriatricianGrade!=="0"){if(!d.geriatricianDate) missing.push("Geriatrician date");if(!d.geriatricianTime) missing.push("Geriatrician time");}
    if(d.postop4AT && d.postop4AT!=="0"){["postop4ATAlertness","postop4ATAMT4","postop4ATAttention","postop4ATAcuteChange"].forEach(k=>{if(!String(d[k]??"").trim()) missing.push(k);});}
  }
  const alive=d.wardDischargeDestination && !["8","8a"].includes(d.wardDischargeDestination) && d.trustDischargeDestination && !["8","8a","8b"].includes(d.trustDischargeDestination);
  if(d.wardDischargeDate && !d.wardDischargeDestination) missing.push("Ward discharge destination");
  if(d.trustDischargeDate && !d.trustDischargeDestination) missing.push("Trust discharge destination");
  if(alive){["residentialStatus","postFractureMobility","medicationAtFollowup"].forEach(k=>{if(!String(d[k]??"").trim()) missing.push(k);});}
  return [...new Set(missing)];
};

function KPI({icon:Icon,label,value,tone="blue"}){const colors={blue:"bg-blue-50 text-blue-700",red:"bg-red-50 text-red-700",amber:"bg-amber-50 text-amber-700",green:"bg-emerald-50 text-emerald-700",violet:"bg-violet-50 text-violet-700"};return <Card className="rounded-2xl border-slate-200 shadow-sm"><CardContent className="flex items-center gap-3 p-4"><div className={`rounded-xl p-2.5 ${colors[tone]}`}><Icon className="h-5 w-5"/></div><div><div className="text-2xl font-extrabold">{value}</div><div className="text-xs font-semibold text-slate-500">{label}</div></div></CardContent></Card>}
function Progress({label,value,color="bg-blue-600"}){return <div><div className="mb-1 flex justify-between text-xs font-semibold"><span>{label}</span><span>{value}%</span></div><div className="h-2 rounded-full bg-slate-200"><div className={`h-2 rounded-full ${color}`} style={{width:`${value}%`}}/></div></div>}

const parseNHFDDateTime=(date,time="00:00")=>{
  const m=String(date||"").match(/^(\d{2})\/(\d{2})\/(\d{4})$/); if(!m) return null;
  const tm=String(time||"00:00").match(/^(\d{2}):(\d{2})$/); if(!tm) return new Date(Number(m[3]),Number(m[2])-1,Number(m[1]));
  const d=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(tm[1]),Number(tm[2]));
  return Number.isNaN(d.getTime())?null:d;
};
const hoursBetween=(aDate,aTime,bDate,bTime)=>{
  const a=parseNHFDDateTime(aDate,aTime),b=parseNHFDDateTime(bDate,bTime); if(!a||!b) return null;
  return (b-a)/3600000;
};
const postOp4ATScore=(d={})=>{
  const vals=[d.postop4ATAlertness,d.postop4ATAMT4,d.postop4ATAttention,d.postop4ATAcuteChange];
  if(vals.some(v=>String(v??"")==="")) return null;
  return Number(vals[0])+Number(vals[1])+Number(vals[2])+Number(vals[3]);
};
const nhfdPerformance=(patients,metric)=>{
  const data=patients.map(p=>p.nhfdData||{});
  const yes=(v)=>v!==undefined&&v!==null&&String(v)!=="";
  const result=(eligible,passed)=>({eligible,passed,percentage:eligible?Math.round((passed/eligible)*100):null});
  if(metric==="ward4h"){
    const eligible=data.filter(d=>yes(d.presentationDate)&&yes(d.wardAdmissionDate)&&d.presentationViaED!=="2"&&d.presentationViaED!=="3"&&d.presentationViaED!=="4");
    return result(eligible.length,eligible.filter(d=>{const h=hoursBetween(d.presentationDate,d.presentationTime,d.wardAdmissionDate,d.wardAdmissionTime);return h!==null&&h>=0&&h<=4;}).length);
  }
  if(metric==="pre4at") return result(data.length,data.filter(d=>d.preop4AT==="2").length);
  if(metric==="geriatric"){
    const eligible=data.filter(d=>yes(d.presentationDate)&&yes(d.geriatricianDate)&&["1","2","4"].includes(d.geriatricianGrade));
    return result(data.length,data.filter(d=>{const h=hoursBetween(d.presentationDate,d.presentationTime,d.geriatricianDate,d.geriatricianTime);return h!==null&&h>=0&&h<=72&&["1","2","4"].includes(d.geriatricianGrade);}).length);
  }
  if(metric==="physio"){
    const eligible=data.filter(d=>d.operated==="1"&&yes(d.surgeryDate));
    return result(eligible.length,eligible.filter(d=>d.physioAssessment==="1").length);
  }
  if(metric==="mobilised"){
    const eligible=data.filter(d=>d.operated==="1"&&yes(d.surgeryDate));
    return result(eligible.length,eligible.filter(d=>["1","2"].includes(d.mobilisedPostSurgery)).length);
  }
  if(metric==="nutrition") return result(data.length,data.filter(d=>["1","2","3"].includes(d.nutritionRisk)).length);
  if(metric==="delirium"){
    const eligible=data.filter(d=>d.operated==="1"&&postOp4ATScore(d)!==null);
    return result(eligible.length,eligible.filter(d=>postOp4ATScore(d)<=3).length);
  }
  if(metric==="falls") return result(data.length,data.filter(d=>d.fallsAssessment==="1").length);
  if(metric==="bone") return result(data.length,data.filter(d=>yes(d.medicationPostFracture)&&d.medicationPostFracture!=="0").length);
  if(metric==="bpt"){
    const eligible=data.filter(d=>d.operated==="1");
    const passed=eligible.filter(d=>{
      const surgeryH=hoursBetween(d.presentationDate,d.presentationTime,d.surgeryDate,d.surgeryTime);
      const gerH=hoursBetween(d.presentationDate,d.presentationTime,d.geriatricianDate,d.geriatricianTime);
      const pre=calculate4ATScore(d);
      const post=postOp4ATScore(d);
      return yes(d.nhsNumber)&&surgeryH!==null&&surgeryH>0&&surgeryH<=36&&gerH!==null&&gerH>=0&&gerH<=72&&["1","2","4"].includes(d.geriatricianGrade)&&d.fallsAssessment==="1"&&yes(d.medicationPostFracture)&&d.medicationPostFracture!=="0"&&pre!==null&&post!==null&&post>=0&&post<=12&&["1","2","3"].includes(d.nutritionRisk)&&d.physioAssessment==="1";
    }).length;
    return result(eligible.length,passed);
  }
  return result(0,0);
};
const nhfdPeriodStart=(period,now=new Date())=>{
  const d=new Date(now);
  if(period==="24h") return new Date(d.getTime()-24*3600000);
  if(period==="7d") return new Date(d.getTime()-7*24*3600000);
  if(period==="30d") return new Date(d.getTime()-30*24*3600000);
  if(period==="12m") return new Date(d.getTime()-365*24*3600000);
  return null;
};
const nhfdPeriodLabel={"24h":"Last 24 hours","7d":"Last 1 week","30d":"Last 30 days","12m":"Last 12 months"};
const nhfdRunChartMonths=(count=12,now=new Date())=>{
  const out=[];
  for(let i=count-1;i>=0;i--){
    const d=new Date(now.getFullYear(),now.getMonth()-i,1);
    out.push({year:d.getFullYear(),month:d.getMonth(),label:d.toLocaleDateString("en-GB",{month:"short",year:"numeric"})});
  }
  return out;
};
const nhfdMonthlyRunData=(patients,metric,count=12,now=new Date())=>nhfdRunChartMonths(count,now).map(m=>{
  const cohort=patients.filter(p=>{
    const d=parseNHFDDateTime(p.nhfdData?.trustDischargeDate);
    return d && d.getFullYear()===m.year && d.getMonth()===m.month;
  });
  const r=nhfdPerformance(cohort,metric);
  return {...m,percentage:r.percentage,passed:r.passed,eligible:r.eligible,cases:cohort.length};
});
const NHFD_RUN_METRICS=[
  ["ward4h","Ward admission ≤4 hours"],["pre4at","Admission 4AT"],["geriatric","Geriatrician ≤72 hours"],["physio","Physiotherapy by day after surgery"],["mobilised","Mobilised by day after surgery"],["nutrition","Nutrition assessment"],["delirium","Not delirious post-op"],["falls","Falls assessment"],["bone","Bone health"],["bpt","Best-practice criteria"]
];
const NHFD_RUN_COLORS=["#2563eb","#7c3aed","#059669","#0891b2","#ea580c","#ca8a04","#dc2626","#db2777","#0f766e","#111827"];

function SectionTitle({title,subtitle}){return <div><h2 className="text-xl font-extrabold tracking-tight">{title}</h2>{subtitle&&<p className="text-sm text-slate-500">{subtitle}</p>}</div>}

export default function App(){
 const [screen,setScreen]=useState("Command Centre");
 const [selectedId,setSelectedId]=useState("UAT-001");
 const [query,setQuery]=useState("");
 const [priority,setPriority]=useState("All");
 const [ward,setWard]=useState("All wards");
 const [mobileOpen,setMobileOpen]=useState(false);
 const [patientRecords,setPatientRecords]=useState(patients);
 const emptyPatient={id:"",name:"",nhs:"",hospital:"",dob:"",age:"",ward:"",bed:"",pathway:"General trauma",diagnosis:"",priority:"Amber",status:"Awaiting review",theatre:"Not listed",procedure:"Pending",fourAT:null,amts:null,og:false,bone:false,falls:false,nhfd:"N/A",owner:"",due:"",action:"",plan:""};
 const [editorOpen,setEditorOpen]=useState(false);
 const [editorMode,setEditorMode]=useState("add");
 const [draft,setDraft]=useState(emptyPatient);
 const openAddPatient=()=>{setEditorMode("add");setDraft({...emptyPatient,id:`UAT-${String(patientRecords.length+1).padStart(3,"0")}`});setEditorOpen(true)};
 const openEditPatient=p=>{setEditorMode("edit");setDraft({...p});setEditorOpen(true)};
 const updateDraft=(field,value)=>setDraft(current=>({...current,[field]:value}));
 const savePatient=()=>{
   if(!draft.name.trim()||!draft.nhs.trim()||!draft.dob.trim()||!draft.ward.trim()||!draft.diagnosis.trim()) return;
   const duplicate=patientRecords.some(p=>p.nhs===draft.nhs&&p.id!==draft.id);
   if(duplicate){alert("A patient with this NHS number already exists.");return;}
   const normalised={...draft,age:Number(draft.age)||0,fourAT:draft.fourAT===""?null:Number(draft.fourAT),amts:draft.amts===""?null:Number(draft.amts)};
   setPatientRecords(records=>editorMode==="add"?[...records,normalised]:records.map(p=>p.id===normalised.id?normalised:p));
   setSelectedId(normalised.id);setEditorOpen(false);setScreen("Patient Detail");
 };
 const [handoverIds,setHandoverIds]=useState(["UAT-001","UAT-002","UAT-006"]);
 const selectedForHandover=patientRecords.filter(p=>handoverIds.includes(p.id));
 const toggleHandover=id=>setHandoverIds(current=>current.includes(id)?current.filter(x=>x!==id):[...current,id]);
 const selectFiltered=()=>setHandoverIds(current=>[...new Set([...current,...filtered.map(p=>p.id)])]);
 const clearHandover=()=>setHandoverIds([]);
 const [quickActionOpen,setQuickActionOpen]=useState(false);
 const [quickActionType,setQuickActionType]=useState("action");
 const [quickActionPatientId,setQuickActionPatientId]=useState(null);
 const [quickActionDraft,setQuickActionDraft]=useState({action:"",owner:"",due:"",status:"",theatre:""});
 const openQuickAction=(type,p)=>{
   setQuickActionType(type);
   setQuickActionPatientId(p.id);
   setQuickActionDraft({action:p.action||"",owner:p.owner||"",due:p.due||"",status:p.status||"",theatre:p.theatre||""});
   setQuickActionOpen(true);
 };
 const saveQuickAction=()=>{
   if(!quickActionPatientId) return;
   setPatientRecords(records=>records.map(p=>p.id===quickActionPatientId?{...p,
     action:quickActionType==="action"?quickActionDraft.action:p.action,
     owner:quickActionType==="action"?quickActionDraft.owner:p.owner,
     due:quickActionType==="action"?quickActionDraft.due:p.due,
     status:quickActionType==="theatre"?quickActionDraft.status:p.status,
     theatre:quickActionType==="theatre"?quickActionDraft.theatre:p.theatre
   }:p));
   setQuickActionOpen(false);
 };
 const selectForHandoverAndStay=id=>{setHandoverIds(current=>current.includes(id)?current:([...current,id]));};
 const printSelectedPatients=()=>{
   const records=selectedForHandover.length?selectedForHandover:(selected?[selected]:[]);
   if(!records.length) return;
   const win=window.open("","_blank","width=1200,height=800");
   if(!win){alert("Please allow pop-ups to print the handover list.");return;}
   const rows=records.map(p=>`<tr><td>${p.name}</td><td>${p.nhs}</td><td>${p.ward} / ${p.bed}</td><td>${p.pathway}</td><td>${p.priority}</td><td>${p.diagnosis}</td><td>${p.status}</td><td>${p.theatre}</td><td>${p.action||""}</td></tr>`).join("");
   win.document.write(`<!doctype html><html><head><title>Trauma & Orthopaedics Handover</title><style>body{font-family:Arial,sans-serif;margin:28px;color:#111}h1{font-size:22px;margin:0 0 4px}p{color:#555;font-size:12px}table{border-collapse:collapse;width:100%;font-size:11px;margin-top:18px}th,td{border:1px solid #cbd5e1;padding:7px;text-align:left;vertical-align:top}th{background:#e2e8f0} .footer{margin-top:20px;font-size:10px;color:#64748b}</style></head><body><h1>Trauma & Orthopaedics Digital Handover</h1><p>Worcestershire Acute Hospitals NHS Trust · Synthetic UAT handover list · ${new Date().toLocaleString()}</p><table><thead><tr><th>Patient</th><th>NHS</th><th>Location</th><th>Pathway</th><th>Priority</th><th>Diagnosis</th><th>Status</th><th>Theatre</th><th>Outstanding action</th></tr></thead><tbody>${rows}</tbody></table><div class="footer">Demonstration only. All records are fictional synthetic UAT data.</div><script>window.onload=()=>window.print();</script></body></html>`);
   win.document.close();
 };
 const selected=patientRecords.find(p=>p.id===selectedId)||patientRecords[0];
 const nof=patientRecords.filter(p=>p.pathway==="NOF");
 const [nofSelectedIds,setNofSelectedIds]=useState([]);
 const nofSelected=nof.filter(p=>nofSelectedIds.includes(p.id));
 const [nofPerformancePeriod,setNofPerformancePeriod]=useState("30d");
 const [nofRunMetric,setNofRunMetric]=useState("bpt");
 const toggleNofSelection=id=>setNofSelectedIds(current=>current.includes(id)?current.filter(x=>x!==id):[...current,id]);
 const selectAllNof=()=>setNofSelectedIds(nof.map(p=>p.id));
 const clearNofSelection=()=>setNofSelectedIds([]);
 const [nofFormOpen,setNofFormOpen]=useState(false);
 const [nofFormMode,setNofFormMode]=useState("add");
 const [nofDraft,setNofDraft]=useState(NHFD_DEFAULTS);
 const openAddNofPatient=()=>{setNofFormMode("add");setNofDraft({...NHFD_DEFAULTS,operatingHospitalCode:"RWP"});setNofFormOpen(true)};
 const openEditNofPatient=p=>{setNofFormMode("edit");setNofDraft({...NHFD_DEFAULTS,...(p.nhfdData||{}),_patientId:p.id});setNofFormOpen(true)};
 const updateNofDraft=(field,value)=>setNofDraft(current=>({...current,[field]:value}));
 const saveNofPatient=()=>{
   const d={...nofDraft,age:calculateAgeFromDob(nofDraft.dob)};
   const missing=nhfdValidation(d);
   if(!d.firstName.trim()||!d.surname.trim()||!d.nhsNumber.trim()||!d.dob.trim()||!d.hospitalNumber.trim()){alert("Please complete the core patient identity fields: first name, surname, NHS number, date of birth and hospital number.");return;}
   const id=nofFormMode==="edit"?nofDraft._patientId:`UAT-NOF-${String(patientRecords.filter(p=>p.pathway==="NOF").length+1).padStart(3,"0")}`;
   const existing=nofFormMode==="edit"?patientRecords.find(p=>p.id===id):null;
   const name=`${d.surname}, ${d.firstName}`;
   const base={id,name,nhs:d.nhsNumber,hospital:d.hospitalNumber,dob:d.dob,age:Number(d.age)||0,ward:d.ward||existing?.ward||"",bed:d.bed||existing?.bed||"",pathway:"NOF",diagnosis:d.diagnosis||existing?.diagnosis||"Hip/femoral fracture",priority:d.priority||existing?.priority||"Amber",status:d.status||existing?.status||(d.operated==="1"?"Post-operative":"Awaiting theatre"),theatre:d.theatre||existing?.theatre||"Not listed",procedure:d.procedure||existing?.procedure||"Pending",fourAT:calculate4ATScore(d),amts:null,og:Boolean(d.geriatricianGrade&&d.geriatricianGrade!=="0"),bone:Boolean(d.medicationPostFracture&&d.medicationPostFracture!=="0"),falls:d.fallsAssessment==="1",nhfdData:d,nhfd:missing.length?"Incomplete":"Ready",owner:d.owner||existing?.owner||"NOF team",due:d.due||existing?.due||"",action:d.action||existing?.action||"",plan:d.plan||existing?.plan||""};
   setPatientRecords(records=>nofFormMode==="edit"?records.map(p=>p.id===id?{...p,...base}:p):[...records,base]);
   setSelectedId(id);setNofFormOpen(false);setScreen("NOF Dashboard");
 };
 const exportNHFD=()=>{
   if(!nofSelected.length){alert("Please select at least one NOF patient.");return;}
   const rows=nofSelected.map(nhfdRow);
   const validation=nofSelected.map((p,i)=>({Patient:p.name,"NHFD v16 validation":p.nhfdData?(nhfdValidation(p.nhfdData).join("; ")||"Ready for export"): (nhfdMissing(rows[i]).join("; ")||"Legacy record - add NHFD data before upload")}));
   const wb=XLSX.utils.book_new();
   const ws=XLSX.utils.aoa_to_sheet([NHFD_V16_HEADERS,...rows]);
   ws["!freeze"]={xSplit:0,ySplit:1};
   ws["!cols"]=NHFD_V16_HEADERS.map(()=>({wch:20}));
   XLSX.utils.book_append_sheet(wb,ws,"NHFD v16 Import");
   const vs=XLSX.utils.json_to_sheet(validation);
   vs["!cols"]=[{wch:28},{wch:80}];
   XLSX.utils.book_append_sheet(wb,vs,"Validation");
   const local=nofSelected.map(p=>({ID:p.id,Patient:p.name,NHS:p.nhs,"Hospital number":p.hospital,DOB:p.dob,Age:p.age,Ward:p.ward,Bed:p.bed,Diagnosis:p.diagnosis,Priority:p.priority,Status:p.status,Procedure:p.procedure,"4AT":p.fourAT,"Orthogeriatrics":p.og,"Bone health":p.bone,"Falls":p.falls,NHFD:p.nhfd}));
   XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(local),"Local source data");
   XLSX.writeFile(wb,`NHFD-v16-selected-${new Date().toISOString().slice(0,10)}.xlsx`);
   const csv=XLSX.utils.sheet_to_csv(ws);
   const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
   const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download=`NHFD-v16-selected-${new Date().toISOString().slice(0,10)}.csv`; a.click(); URL.revokeObjectURL(url);
 };
 const filtered=useMemo(()=>patientRecords.filter(p=>{
   const q=query.toLowerCase();
   return `${p.name} ${p.nhs} ${p.hospital} ${p.diagnosis} ${p.ward}`.toLowerCase().includes(q)
    &&(priority==="All"||p.priority===priority)&&(ward==="All wards"||ward==="All"||p.ward===ward);
 }),[patientRecords,query,priority,ward]);
 const openPatient=p=>{setSelectedId(p.id);setScreen("Patient Detail")};
 const red=patientRecords.filter(p=>p.priority==="Red");
 const theatre=patientRecords.filter(p=>["Theatre 1","Theatre 2","Backup"].includes(p.theatre));
 const openActions=patientRecords.filter(p=>p.action);
 const wardCounts=[...new Set(patientRecords.map(p=>p.ward))].map(w=>[w,patientRecords.filter(p=>p.ward===w).length]);

 const Sidebar=()=> <aside className={`fixed inset-y-0 left-0 z-40 w-64 bg-slate-950 text-white transition-transform lg:static lg:translate-x-0 ${mobileOpen?"translate-x-0":"-translate-x-full"}`}>
   <div className="flex h-20 items-center gap-3 border-b border-white/10 px-5"><div className="rounded-xl bg-[#005eb8] p-2"><Hospital className="h-6 w-6"/></div><div><div className="font-bold">WAHT Trauma</div><div className="text-xs text-slate-400">Synthetic UAT prototype</div></div><button className="ml-auto lg:hidden" onClick={()=>setMobileOpen(false)}><X/></button></div>
   <nav className="space-y-1 p-3">{nav.map(([label,Icon])=><button key={label} onClick={()=>{setScreen(label);setMobileOpen(false)}} className={`flex w-full items-center gap-3 rounded-xl px-3 py-3 text-left text-sm font-semibold transition ${screen===label?"bg-[#005eb8] text-white shadow-lg":"text-slate-300 hover:bg-white/10 hover:text-white"}`}><Icon className="h-5 w-5"/>{label}</button>)}</nav>
   <div className="absolute bottom-0 w-full border-t border-white/10 p-4"><div className="rounded-xl bg-white/10 p-3 text-xs text-slate-300"><ShieldCheck className="mb-2 h-5 w-5 text-emerald-400"/><b>Demonstration only</b><br/>All records are fictional synthetic UAT data.</div></div>
 </aside>;

 const PatientMiniCard=({p})=><button onClick={()=>openPatient(p)} className="w-full rounded-xl border border-slate-200 bg-white p-3 text-left transition hover:border-blue-300 hover:shadow-md"><div className="flex items-start justify-between gap-2"><div><div className="font-bold">{p.name}</div><div className="text-xs text-slate-500">{p.ward} · Bed {p.bed}</div></div><Badge className={`${pClass[p.priority]} border`}>{p.priority}</Badge></div><div className="mt-2 text-sm">{p.diagnosis}</div><div className="mt-2 flex items-center justify-between text-xs text-slate-500"><span>{p.status}</span><ChevronRight className="h-4 w-4"/></div></button>;

 const CommandCentre=()=> <div className="space-y-5">
   <div className="flex flex-col justify-between gap-3 md:flex-row md:items-end"><SectionTitle title="Trauma Command Centre" subtitle="Large-screen operational view using the 20 synthetic UAT records"/><div className="rounded-full bg-emerald-100 px-3 py-1 text-xs font-bold text-emerald-800">● LIVE DEMO · Updated 08:30</div></div>
   <div className="grid grid-cols-2 gap-3 md:grid-cols-3 xl:grid-cols-6"><KPI icon={BedDouble} label="Active patients" value={patientRecords.length}/><KPI icon={HeartPulse} label="NOF patients" value={nof.length} tone="violet"/><KPI icon={AlertTriangle} label="Red alerts" value={red.length} tone="red"/><KPI icon={Syringe} label="Planned/backup" value={theatre.length} tone="amber"/><KPI icon={ClipboardCheck} label="Open actions" value={openActions.length} tone="blue"/><KPI icon={CheckCircle2} label="NHFD ready" value={nof.filter(p=>p.nhfd==="Ready").length} tone="green"/></div>
   <div className="grid gap-5 xl:grid-cols-3">
     <Card className="rounded-2xl xl:col-span-1"><CardHeader><CardTitle className="flex items-center gap-2 text-base"><AlertTriangle className="h-5 w-5 text-red-600"/>Priority alerts</CardTitle></CardHeader><CardContent className="max-h-[430px] space-y-2 overflow-auto">{red.map(p=><PatientMiniCard key={p.id} p={p}/>)}</CardContent></Card>
     <Card className="rounded-2xl xl:col-span-1"><CardHeader><CardTitle className="flex items-center gap-2 text-base"><Theater className="h-5 w-5 text-violet-600"/>Theatre board</CardTitle></CardHeader><CardContent className="space-y-4">{["Theatre 1","Theatre 2","Backup"].map(t=><div key={t}><div className="mb-2 flex items-center justify-between"><b className="text-sm">{t}</b><Badge variant="outline">{patientRecords.filter(p=>p.theatre===t).length}</Badge></div><div className="space-y-2">{patientRecords.filter(p=>p.theatre===t).map((p,i)=><button key={p.id} onClick={()=>openPatient(p)} className="flex w-full items-center gap-3 rounded-xl bg-slate-50 p-3 text-left hover:bg-blue-50"><span className="flex h-7 w-7 items-center justify-center rounded-lg bg-[#005eb8] text-xs font-bold text-white">{i+1}</span><span className="min-w-0 flex-1"><b className="block truncate text-sm">{p.name}</b><span className="block truncate text-xs text-slate-500">{p.procedure}</span></span><Badge className={`${pClass[p.priority]} border`}>{p.priority}</Badge></button>)}</div></div>)}</CardContent></Card>
     <div className="space-y-5"><Card className="rounded-2xl"><CardHeader><CardTitle className="text-base">NOF best-practice snapshot</CardTitle></CardHeader><CardContent className="space-y-4"><Progress label="4AT recorded" value={pct(nof.filter(p=>p.fourAT!==null).length,nof.length)}/><Progress label="Orthogeriatric complete" value={pct(nof.filter(p=>p.og).length,nof.length)} color="bg-violet-600"/><Progress label="Bone health complete" value={pct(nof.filter(p=>p.bone).length,nof.length)} color="bg-emerald-600"/><Progress label="Falls assessment complete" value={pct(nof.filter(p=>p.falls).length,nof.length)} color="bg-amber-500"/></CardContent></Card><Card className="rounded-2xl"><CardHeader><CardTitle className="text-base">Ward distribution</CardTitle></CardHeader><CardContent className="space-y-2">{wardCounts.map(([w,n])=><button key={w} onClick={()=>{setWard(w);setScreen("Patients Admission List")}} className="flex w-full items-center justify-between rounded-xl bg-slate-50 px-3 py-2 text-sm hover:bg-blue-50"><span>{w}</span><Badge>{n}</Badge></button>)}</CardContent></Card></div>
   </div>
 </div>;

 const TraumaList=()=> <div className="space-y-5">
   <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
     <SectionTitle title="Trauma Patient List" subtitle="Search, filter and tick patients for the printable handover list"/>
     <div className="flex flex-wrap items-center gap-2">
       <Button onClick={openAddPatient} className="bg-[#005eb8] hover:bg-[#004b93]"> <Plus className="mr-2 h-4 w-4"/>Add New Patient</Button>
       <Badge className="bg-blue-100 px-3 py-2 text-blue-800">{handoverIds.length} selected</Badge>
       <Button variant="outline" onClick={selectFiltered}><CheckCircle2 className="mr-2 h-4 w-4"/>Select filtered</Button>
       <Button variant="outline" onClick={clearHandover}>Clear selection</Button>
       <Button disabled={!handoverIds.length} onClick={()=>setScreen("MDT / Handover")} className="bg-[#005eb8]"><ClipboardCheck className="mr-2 h-4 w-4"/>Open handover</Button>
       <Button disabled={!handoverIds.length} onClick={printSelectedPatients} className="bg-slate-950 hover:bg-slate-800"><Printer className="mr-2 h-4 w-4"/>Export Selected Patients PDF/Print</Button>
     </div>
   </div>
   <Card className="rounded-2xl"><CardContent className="grid gap-4 p-4 md:grid-cols-[1fr_auto_180px]"><label className="relative"><Search className="absolute left-3 top-3 h-4 w-4 text-slate-400"/><Input className="pl-9" placeholder="Search name, NHS number, ward or diagnosis" value={query} onChange={e=>setQuery(e.target.value)}/></label><div className="flex flex-wrap items-center gap-2" aria-label="Priority filter"><span className="mr-1 text-xs font-bold uppercase tracking-wide text-slate-500">Priority</span>{[["All","bg-slate-900 text-white border-slate-900","bg-slate-100 text-slate-700 border-slate-200"],["Red","bg-red-600 text-white border-red-600","bg-red-50 text-red-700 border-red-200"],["Amber","bg-amber-500 text-white border-amber-500","bg-amber-50 text-amber-700 border-amber-200"],["Green","bg-emerald-600 text-white border-emerald-600","bg-emerald-50 text-emerald-700 border-emerald-200"]].map(([label,active,inactive])=><button key={label} type="button" onClick={()=>setPriority(label)} aria-pressed={priority===label} className={`rounded-full border px-4 py-2 text-sm font-bold transition ${priority===label?active:inactive} hover:shadow-sm`}>{label}</button>)}</div><select className="h-10 rounded-md border bg-white px-3 text-sm" value={ward} onChange={e=>setWard(e.target.value)}><option>All wards</option>{[...new Set(patientRecords.map(p=>p.ward))].map(w=><option key={w}>{w}</option>)}</select></CardContent></Card>
   <Card className="overflow-hidden rounded-2xl"><CardContent className="p-0"><div className="overflow-x-auto"><table className="w-full min-w-[1160px] text-left text-sm"><thead className="bg-slate-100 text-xs uppercase tracking-wide text-slate-500"><tr><th className="px-4 py-3">Print</th><th className="px-4 py-3">Priority</th><th className="px-4 py-3">Patient</th><th className="px-4 py-3">Identifiers</th><th className="px-4 py-3">Location</th><th className="px-4 py-3">Pathway</th><th className="px-4 py-3">Diagnosis</th><th className="px-4 py-3">Status</th><th className="px-4 py-3 text-right">Actions</th></tr></thead><tbody className="divide-y">{filtered.map(p=><tr key={p.id} className={`${handoverIds.includes(p.id)?"bg-blue-50":"bg-white"} hover:bg-blue-50`}><td className="px-4 py-3"><input aria-label={`Select ${p.name} for handover`} type="checkbox" checked={handoverIds.includes(p.id)} onChange={()=>toggleHandover(p.id)} className="h-5 w-5 cursor-pointer accent-[#005eb8]"/></td><td className="px-4 py-3"><Badge className={`${pClass[p.priority]} border`}>{p.priority}</Badge></td><td className="cursor-pointer px-4 py-3" onClick={()=>openPatient(p)}><b>{p.name}</b><div className="text-xs text-slate-500">Age {p.age} · {p.id}</div></td><td className="px-4 py-3"><span className="block">NHS {p.nhs}</span><span className="text-xs text-slate-500">Hosp {p.hospital} · DOB {p.dob}</span></td><td className="px-4 py-3">{p.ward}<div className="text-xs text-slate-500">Bed {p.bed}</div></td><td className="px-4 py-3"><Badge variant="outline">{p.pathway}</Badge></td><td className="max-w-[280px] px-4 py-3 font-medium">{p.diagnosis}</td><td className="px-4 py-3">{p.status}</td><td className="px-4 py-3"><div className="flex items-center justify-end gap-1"><Button variant="outline" size="sm" onClick={(e)=>{e.stopPropagation();openEditPatient(p)}} className="h-9 px-2" title={`Edit ${p.name}`} aria-label={`Edit ${p.name}`}><Pencil className="h-4 w-4"/><span className="sr-only">Edit</span></Button><button type="button" className="rounded-lg p-2 hover:bg-slate-100" onClick={()=>openPatient(p)} aria-label={`Open ${p.name}`}><ChevronRight className="h-5 w-5 text-slate-400"/></button></div></td></tr>)}</tbody></table></div></CardContent></Card>
   {handoverIds.length>0&&<div className="sticky bottom-4 flex items-center justify-between rounded-2xl bg-slate-950 p-4 text-white shadow-2xl"><div><b>{handoverIds.length} patient{handoverIds.length===1?"":"s"} selected</b><div className="text-xs text-slate-300">Ready for handover review and printing</div></div><Button onClick={()=>setScreen("MDT / Handover")} className="bg-[#005eb8]">Continue to handover</Button></div>}
 </div>;
 const QuickActionModal=()=> quickActionOpen&&<div className="fixed inset-0 z-[80] flex items-center justify-center bg-slate-950/70 p-4" onMouseDown={e=>{if(e.target===e.currentTarget)setQuickActionOpen(false)}}><Card className="w-full max-w-2xl rounded-2xl shadow-2xl"><CardHeader className="flex flex-row items-center justify-between border-b"><div><CardTitle>{quickActionType==="action"?"Add action":"Theatre update"}</CardTitle><p className="mt-1 text-sm text-slate-500">Update the selected patient's handover record.</p></div><button aria-label="Close quick action" onClick={()=>setQuickActionOpen(false)} className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5"/></button></CardHeader><CardContent className="space-y-4 p-5">{quickActionType==="action"?<><label className="block text-sm font-semibold">Outstanding action<textarea autoFocus className="mt-1 min-h-28 w-full rounded-md border p-3" value={quickActionDraft.action} onChange={e=>setQuickActionDraft(v=>({...v,action:e.target.value}))}/></label><div className="grid gap-4 sm:grid-cols-2"><label className="text-sm font-semibold">Owner<Input value={quickActionDraft.owner} onChange={e=>setQuickActionDraft(v=>({...v,owner:e.target.value}))}/></label><label className="text-sm font-semibold">Due time<Input type="time" value={quickActionDraft.due} onChange={e=>setQuickActionDraft(v=>({...v,due:e.target.value}))}/></label></div></>:<div className="grid gap-4 sm:grid-cols-2"><label className="text-sm font-semibold">Theatre<select className="mt-1 h-10 w-full rounded-md border bg-white px-3" value={quickActionDraft.theatre} onChange={e=>setQuickActionDraft(v=>({...v,theatre:e.target.value}))}><option>Theatre 1</option><option>Theatre 2</option><option>Backup</option><option>Not listed</option><option>Tertiary centre</option><option>Possible</option></select></label><label className="text-sm font-semibold">Status<Input value={quickActionDraft.status} onChange={e=>setQuickActionDraft(v=>({...v,status:e.target.value}))}/></label></div>}<div className="flex justify-end gap-2 border-t pt-4"><Button variant="outline" onClick={()=>setQuickActionOpen(false)}>Cancel</Button><Button className="bg-[#005eb8]" onClick={saveQuickAction}>Save update</Button></div></CardContent></Card></div>;

 const PatientEditor=()=> editorOpen&&<div className="fixed inset-0 z-[70] flex items-center justify-center bg-slate-950/70 p-4" onMouseDown={e=>{if(e.target===e.currentTarget)setEditorOpen(false)}}><Card className="max-h-[92vh] w-full max-w-5xl overflow-y-auto rounded-2xl shadow-2xl"><CardHeader className="sticky top-0 z-10 flex flex-row items-center justify-between border-b bg-white"><div><CardTitle>{editorMode==="add"?"Add New Patient":"Edit Patient Entry"}</CardTitle><p className="mt-1 text-sm text-slate-500">Required identity and clinical handover fields are marked *</p></div><button aria-label="Close patient form" onClick={()=>setEditorOpen(false)} className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5"/></button></CardHeader><CardContent className="space-y-6 p-6"><div><h3 className="mb-3 font-bold text-[#005eb8]">Patient identity</h3><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3">{[["name","Full name *","SURNAME, First name"],["nhs","NHS number *","10 digits"],["hospital","Hospital number","Hospital identifier"],["dob","Date of birth *","DD/MM/YYYY"],["age","Age","Calculated age"],["id","Record ID","UAT identifier"]].map(([f,l,ph])=><label key={f} className="space-y-1 text-sm font-semibold">{l}<Input disabled={f==="id"} value={draft[f]??""} placeholder={ph} onChange={e=>updateDraft(f,e.target.value)}/></label>)}</div></div><div><h3 className="mb-3 font-bold text-[#005eb8]">Location and pathway</h3><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><label className="space-y-1 text-sm font-semibold">Ward *<Input value={draft.ward} onChange={e=>updateDraft("ward",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Bed<Input value={draft.bed} onChange={e=>updateDraft("bed",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Pathway<select className="mt-1 h-10 w-full rounded-md border bg-white px-3" value={draft.pathway} onChange={e=>updateDraft("pathway",e.target.value)}><option>General trauma</option><option>NOF</option></select></label><label className="space-y-1 text-sm font-semibold">Priority<select className="mt-1 h-10 w-full rounded-md border bg-white px-3" value={draft.priority} onChange={e=>updateDraft("priority",e.target.value)}><option>Red</option><option>Amber</option><option>Green</option></select></label></div></div><div><h3 className="mb-3 font-bold text-[#005eb8]">Clinical handover</h3><div className="grid gap-4 md:grid-cols-2"><label className="space-y-1 text-sm font-semibold">Diagnosis *<Input value={draft.diagnosis} onChange={e=>updateDraft("diagnosis",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Status<Input value={draft.status} onChange={e=>updateDraft("status",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Theatre<Input value={draft.theatre} onChange={e=>updateDraft("theatre",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Procedure<Input value={draft.procedure} onChange={e=>updateDraft("procedure",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold md:col-span-2">Current plan<textarea className="mt-1 min-h-20 w-full rounded-md border p-3" value={draft.plan} onChange={e=>updateDraft("plan",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold md:col-span-2">Outstanding action<textarea className="mt-1 min-h-20 w-full rounded-md border p-3" value={draft.action} onChange={e=>updateDraft("action",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Action owner<Input value={draft.owner} onChange={e=>updateDraft("owner",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">Due time<Input type="time" value={draft.due} onChange={e=>updateDraft("due",e.target.value)}/></label></div></div>{draft.pathway==="NOF"&&<div className="rounded-2xl border border-violet-200 bg-violet-50 p-4"><h3 className="mb-3 font-bold text-violet-800">NOF pathway fields</h3><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><label className="space-y-1 text-sm font-semibold">4AT<Input type="number" min="0" value={draft.fourAT??""} onChange={e=>updateDraft("fourAT",e.target.value)}/></label><label className="space-y-1 text-sm font-semibold">AMTS<Input type="number" min="0" max="10" value={draft.amts??""} onChange={e=>updateDraft("amts",e.target.value)}/></label>{[["og","Orthogeriatric review"],["bone","Bone health"],["falls","Falls assessment"]].map(([f,l])=><label key={f} className="flex items-center gap-2 pt-6 text-sm font-semibold"><input type="checkbox" checked={Boolean(draft[f])} onChange={e=>updateDraft(f,e.target.checked)} className="h-5 w-5 accent-violet-700"/>{l} complete</label>)}</div></div>}<div className="flex flex-col-reverse justify-end gap-2 border-t pt-4 sm:flex-row"><Button variant="outline" onClick={()=>setEditorOpen(false)}>Cancel</Button><Button onClick={savePatient} disabled={!draft.name.trim()||!draft.nhs.trim()||!draft.dob.trim()||!draft.ward.trim()||!draft.diagnosis.trim()} className="bg-[#005eb8]">{editorMode==="add"?"Add patient to list":"Save changes"}</Button></div></CardContent></Card></div>;


 const PatientDetail=()=> <div className="space-y-5"><div className="flex items-center justify-between"><SectionTitle title="Patient Detail" subtitle="Synthetic record with full-name identity block"/><Button variant="outline" onClick={()=>setScreen("Patients Admission List")}>Back to list</Button></div><div className="grid gap-5 xl:grid-cols-[1fr_380px]"><div className="space-y-5"><Card className="overflow-hidden rounded-2xl"><div className="bg-slate-950 p-6 text-white"><div className="flex flex-col justify-between gap-3 md:flex-row"><div><p className="text-xs font-bold uppercase tracking-[.2em] text-slate-400">{selected.id}</p><h2 className="mt-1 text-3xl font-extrabold">{selected.name}</h2><p className="mt-1 text-slate-300">DOB {selected.dob} · Age {selected.age} · {selected.ward}, bed {selected.bed}</p></div><Badge className={`${pClass[selected.priority]} h-fit border px-4 py-2 text-sm`}>{selected.priority} priority</Badge></div></div><CardContent className="grid gap-4 p-5 md:grid-cols-3">{[["NHS number",selected.nhs],["Hospital number",selected.hospital],["Pathway",selected.pathway],["Diagnosis",selected.diagnosis],["Status",selected.status],["Theatre",selected.theatre]].map(([k,v])=><div key={k} className="rounded-xl bg-slate-50 p-3"><div className="text-xs font-semibold uppercase text-slate-400">{k}</div><div className="mt-1 font-bold">{v}</div></div>)}</CardContent></Card><Card className="rounded-2xl"><CardHeader><CardTitle className="flex items-center gap-2 text-base"><Stethoscope className="h-5 w-5 text-blue-600"/>Current clinical plan</CardTitle></CardHeader><CardContent><p>{selected.plan}</p><div className="mt-4 rounded-xl border border-amber-200 bg-amber-50 p-4"><b>Outstanding action</b><p className="mt-1 text-sm">{selected.action}</p><div className="mt-2 text-xs font-semibold text-amber-800">Owner: {selected.owner} · Due: {selected.due}</div></div></CardContent></Card></div><div className="space-y-5">{isNOF(selected)&&<Card className="rounded-2xl border-violet-200"><CardHeader><CardTitle className="flex items-center gap-2 text-base"><HeartPulse className="h-5 w-5 text-violet-600"/>NOF best-practice</CardTitle></CardHeader><CardContent className="grid grid-cols-2 gap-3">{[["4AT",selected.fourAT],["AMTS",selected.amts],["Orthogeriatrics",selected.og?"Complete":"Outstanding"],["Bone health",selected.bone?"Complete":"Outstanding"],["Falls",selected.falls?"Complete":"Outstanding"],["NHFD",selected.nhfd]].map(([k,v])=><div key={k} className="rounded-xl bg-violet-50 p-3"><div className="text-xs text-violet-600">{k}</div><div className="font-bold text-violet-950">{v}</div></div>)}</CardContent></Card>}<Card className="rounded-2xl"><CardHeader><CardTitle className="text-base">Quick actions</CardTitle></CardHeader><CardContent className="grid gap-2"><Button className="bg-[#005eb8]" onClick={()=>openEditPatient(selected)}>Edit synthetic record</Button><Button variant="outline" onClick={()=>openQuickAction("action",selected)}>Add action</Button><Button variant="outline" onClick={()=>openQuickAction("theatre",selected)}>Theatre update</Button><Button variant="outline" onClick={()=>{selectForHandoverAndStay(selected.id);setScreen("Patients Admission List")}}>Select for handover print</Button><Button variant="outline" onClick={printSelectedPatients}>Export selected patients (PDF/Print)</Button></CardContent></Card></div></div></div>;
 function isNOF(p){return p.pathway==="NOF"}

 const NofPatientForm=()=>{
   if(!nofFormOpen) return null;
   const d=nofDraft; const set=updateNofDraft; const operated=d.operated==="1"; const pre4=d.preop4AT==="2"; const post4=d.postop4AT && d.postop4AT!=="0"; const geriatrician=d.geriatricianGrade && d.geriatricianGrade!=="0";
   const Field=({label,field,placeholder="",type="text",required=false,disabled=false})=><label className="space-y-1 text-sm font-semibold">{label}{required&&<span className="text-red-600"> *</span>}<Input type={type} disabled={disabled} value={d[field]??""} placeholder={placeholder} onChange={e=>set(field,e.target.value)}/></label>;
   const Select=({label,field,options,required=false,placeholder="Select..."})=><label className="space-y-1 text-sm font-semibold">{label}{required&&<span className="text-red-600"> *</span>}<select className="mt-1 h-10 w-full rounded-md border bg-white px-3 text-sm" value={d[field]??""} onChange={e=>set(field,e.target.value)}><option value="">{placeholder}</option>{options.map(([v,l])=><option key={v} value={v}>{v} — {l}</option>)}</select></label>;
   const Multi=({label,field,options})=><div className="text-sm font-semibold"><div className="mb-2">{label}</div><div className="grid gap-2 sm:grid-cols-2">{options.map(([v,l])=><label key={v} className="flex items-center gap-2 rounded-lg border bg-white p-2 font-normal"><input type="checkbox" checked={(d[field]||[]).includes(v)} onChange={e=>set(field,e.target.checked?[...(d[field]||[]),v]:(d[field]||[]).filter(x=>x!==v))} className="h-4 w-4 accent-[#005eb8]"/>{v} — {l}</label>)}</div></div>;
   const Section=({title,children})=><div className="rounded-2xl border border-slate-200 bg-slate-50/70 p-4"><h3 className="mb-4 font-bold text-[#005eb8]">{title}</h3>{children}</div>;
   const missing=nhfdValidation(d);
   return <div className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/70 p-3 md:p-6" onMouseDown={e=>{if(e.target===e.currentTarget)setNofFormOpen(false)}}><Card className="max-h-[95vh] w-full max-w-7xl overflow-y-auto rounded-2xl shadow-2xl"><CardHeader className="sticky top-0 z-20 flex flex-row items-center justify-between border-b bg-white"><div><CardTitle>{nofFormMode==="add"?"Add NOF patient — NHFD v16 data form":"Edit NOF patient — NHFD v16 data"}</CardTitle><p className="mt-1 text-sm text-slate-500">This form mirrors the NHFD v16 (2026) import columns. Calculated 4AT scores are not entered because NHFD calculates them.</p></div><button aria-label="Close NOF form" onClick={()=>setNofFormOpen(false)} className="rounded-lg p-2 hover:bg-slate-100"><X className="h-5 w-5"/></button></CardHeader><CardContent className="space-y-5 p-4 md:p-6">
     <div className="rounded-xl border border-amber-200 bg-amber-50 p-3 text-sm text-amber-950"><b>Export status:</b> {missing.length?`${missing.length} field(s) still need attention.`:"All currently applicable NHFD v16 fields are completed."} <span className="ml-2 text-xs">Optional/conditional fields become relevant as the patient's pathway progresses.</span></div>
     <Section title="Local NOF / dashboard details"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Field label="Ward" field="ward" required/><Field label="Bed" field="bed"/><Field label="Diagnosis" field="diagnosis" required/><Select label="Priority" field="priority" options={[["Red","Red"],["Amber","Amber"],["Green","Green"]]}/><Select label="Status" field="status" options={[["Awaiting theatre","Awaiting theatre"],["Theatre today","Theatre today"],["Theatre planned","Theatre planned"],["Post-operative","Post-operative"],["Discharged","Discharged"]]}/><Field label="Theatre" field="theatre"/><Field label="Procedure" field="procedure"/><Field label="Action owner" field="owner"/><Field label="Due time" field="due" type="time"/><div className="md:col-span-2 lg:col-span-4"><label className="space-y-1 text-sm font-semibold">Current plan<textarea className="mt-1 min-h-20 w-full rounded-md border bg-white p-3" value={d.plan??""} onChange={e=>set("plan",e.target.value)}/></label></div><div className="md:col-span-2 lg:col-span-4"><label className="space-y-1 text-sm font-semibold">Outstanding action<textarea className="mt-1 min-h-20 w-full rounded-md border bg-white p-3" value={d.action??""} onChange={e=>set("action",e.target.value)}/></label></div></div></Section>
     <Section title="1. Patient"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Field label="Operating hospital code" field="operatingHospitalCode" required/><Field label="NHS number" field="nhsNumber" required placeholder="10 digits / Overseas / [NONHS]"/><Field label="Hospital number / patient ID" field="hospitalNumber" required/><Field label="First name" field="firstName" required/><Field label="Surname" field="surname" required/><Field label="Date of birth" field="dob" required placeholder="DD/MM/YYYY"/><Select label="Sex / Gender" field="sex" options={NHFD_OPTIONS.sex} required/><Field label="Patient postcode" field="postcode" placeholder="e.g. WR1 1AA"/></div></Section>
     <Section title="2. Admission"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Field label="First hospital code" field="firstHospitalCode" placeholder="X if not transferred"/><Field label="Presentation date" field="presentationDate" required placeholder="DD/MM/YYYY"/><Field label="Presentation time" field="presentationTime" required placeholder="HH:MM"/><Select label="Ambulance" field="ambulance" options={NHFD_OPTIONS.ambulance}/><Field label="Ambulance date" field="ambulanceDate" placeholder="DD/MM/YYYY"/><Field label="Ambulance time" field="ambulanceTime" placeholder="HH:MM"/><Select label="Residence before admission" field="residenceBeforeAdmission" options={NHFD_OPTIONS.residenceBeforeAdmission} required/><Select label="Presentation via ED" field="presentationViaED" options={NHFD_OPTIONS.presentationViaED} required/><Select label="Admitted to ortho/orthogeriatric ward" field="admittedToOrthoWard" options={NHFD_OPTIONS.admittedToOrthoWard} required/><Field label="Date ward admission" field="wardAdmissionDate" placeholder="DD/MM/YYYY"/><Field label="Time ward admission" field="wardAdmissionTime" placeholder="HH:MM"/><Select label="Nerve block" field="nerveBlockED" options={NHFD_OPTIONS.nerveBlockED} required/></div></Section>
     <Section title="3. Assessment"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Select label="Pre-fracture mobility" field="mobilityPreFracture" options={NHFD_OPTIONS.mobilityPreFracture} required/><Select label="Admission 4AT" field="preop4AT" options={NHFD_OPTIONS.preop4AT} required/><Select label="Pre-op 4AT — Alertness" field="preop4ATAlertness" options={NHFD_OPTIONS.fourATAlertness}/><Select label="Pre-op 4AT — AMT4" field="preop4ATAMT4" options={NHFD_OPTIONS.fourATAMT4}/><Select label="Pre-op 4AT — Attention" field="preop4ATAttention" options={NHFD_OPTIONS.fourATAttention}/><Select label="Pre-op 4AT — Acute change" field="preop4ATAcuteChange" options={NHFD_OPTIONS.fourATAcute}/><Select label="ASA grade" field="asaGrade" options={NHFD_OPTIONS.asaGrade} required/><Select label="Nutrition risk assessment" field="nutritionRisk" options={NHFD_OPTIONS.nutritionRisk} required/><Select label="Bone protection pre-fracture" field="boneProtectionPre" options={NHFD_OPTIONS.boneProtectionPre} required/></div></Section>
     <Section title="4. Fracture"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Select label="Fracture type" field="fractureType" options={NHFD_OPTIONS.fractureType} required/><Select label="Fracture side" field="fractureSide" options={NHFD_OPTIONS.fractureSide} required/><Select label="Atypical/pathological fracture" field="pathologicalFracture" options={NHFD_OPTIONS.pathologicalFracture} required/></div></Section>
     <Section title="5. Surgery"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Select label="Was an operation performed?" field="operated" options={NHFD_OPTIONS.operated} required/><Field label="Surgery date" field="surgeryDate" placeholder="DD/MM/YYYY"/><Field label="Surgery time" field="surgeryTime" placeholder="HH:MM"/><Select label="Type of operation" field="operationType" options={NHFD_OPTIONS.operationType}/><div className="md:col-span-2 lg:col-span-4"><Multi label="Modes of anaesthesia used" field="anaesthesia" options={NHFD_OPTIONS.anaesthesia}/></div><Select label="Delay reason (>36 h)" field="delayReason" options={NHFD_OPTIONS.delayReason}/><Select label="Most senior scrubbed surgeon" field="scrubbedSurgeonGrade" options={NHFD_OPTIONS.clinicianGrade}/><Select label="Most senior surgeon present" field="surgeonGrade" options={NHFD_OPTIONS.clinicianGrade}/><Select label="Most senior anaesthetist" field="anaesthetistGrade" options={NHFD_OPTIONS.anaesthetistGrade}/><Select label="Full weight-bearing documented" field="weightBearPostOp" options={NHFD_OPTIONS.weightBearPostOp}/></div></Section>
     <Section title="6. Ward care"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Select label="Physiotherapist assessment" field="physioAssessment" options={NHFD_OPTIONS.physioAssessment}/><Select label="Mobilised post surgery" field="mobilisedPostSurgery" options={NHFD_OPTIONS.mobilisedPostSurgery}/><Select label="Geriatrician grade" field="geriatricianGrade" options={NHFD_OPTIONS.geriatricianGrade}/><Field label="Geriatrician date" field="geriatricianDate" placeholder="DD/MM/YYYY"/><Field label="Geriatrician time" field="geriatricianTime" placeholder="HH:MM"/><Select label="Specialist falls assessment" field="fallsAssessment" options={NHFD_OPTIONS.fallsAssessment}/><Select label="New pressure ulcer grade 2+" field="pressureUlcers" options={NHFD_OPTIONS.pressureUlcers}/><Select label="Bone protection medication after fracture" field="medicationPostFracture" options={NHFD_OPTIONS.medicationPostFracture}/></div></Section>
     <Section title="7. Delirium"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Select label="Repeat 4AT after operation" field="postop4AT" options={NHFD_OPTIONS.postop4AT}/><Select label="Post-op 4AT — Alertness" field="postop4ATAlertness" options={NHFD_OPTIONS.fourATAlertness}/><Select label="Post-op 4AT — AMT4" field="postop4ATAMT4" options={NHFD_OPTIONS.fourATAMT4}/><Select label="Post-op 4AT — Attention" field="postop4ATAttention" options={NHFD_OPTIONS.fourATAttention}/><Select label="Post-op 4AT — Acute change" field="postop4ATAcuteChange" options={NHFD_OPTIONS.fourATAcute}/></div></Section>
     <Section title="8. Discharge"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4"><Field label="Ward discharge date" field="wardDischargeDate" placeholder="DD/MM/YYYY"/><Select label="Ward discharge destination" field="wardDischargeDestination" options={NHFD_OPTIONS.wardDischargeDestination}/><Field label="Trust discharge date" field="trustDischargeDate" placeholder="DD/MM/YYYY"/><Select label="Trust discharge destination" field="trustDischargeDestination" options={NHFD_OPTIONS.trustDischargeDestination}/></div></Section>
     <Section title="9. Re-operations"><div className="grid gap-4 md:grid-cols-2"><Select label="Re-operations within 120 days" field="reoperations" options={NHFD_OPTIONS.reoperations}/></div></Section>
     <Section title="10. Follow-up"><div className="grid gap-4 md:grid-cols-2 lg:grid-cols-3"><Select label="Residential status" field="residentialStatus" options={NHFD_OPTIONS.residentialStatus}/><Select label="Post-fracture mobility" field="postFractureMobility" options={NHFD_OPTIONS.postFractureMobility}/><Select label="Bone protection medication at follow-up" field="medicationAtFollowup" options={NHFD_OPTIONS.medicationAtFollowup}/></div></Section>
     <div className="rounded-xl bg-slate-100 p-3 text-xs text-slate-600"><b>Calculated fields:</b> NHFD v16 calculates the pre-op and post-op 4AT scores from their component fields. They are deliberately not included in the import CSV.</div>
     <div className="flex flex-col-reverse justify-between gap-3 border-t pt-4 sm:flex-row"><div className="text-xs text-slate-500">{missing.length?`${missing.length} validation item(s) remain.`:"Ready — no current validation gaps detected."}</div><div className="flex gap-2"><Button variant="outline" onClick={()=>setNofFormOpen(false)}>Cancel</Button><Button onClick={saveNofPatient} className="bg-[#005eb8]"><Save className="mr-2 h-4 w-4"/>{nofFormMode==="add"?"Save NOF patient":"Save NHFD data"}</Button></div></div>
   </CardContent></Card></div>;
 };

 const NofDashboard=()=>{
   const periodStart=nhfdPeriodStart(nofPerformancePeriod);
   const periodPatients=nof.filter(p=>{
     const d=parseNHFDDateTime(p.nhfdData?.presentationDate,p.nhfdData?.presentationTime);
     return d && (!periodStart||d>=periodStart);
   });
   const periodLabels=nhfdPeriodLabel;
   const metrics=[
     ["ward4h","Admission to orthopaedic ward within 4 hours"],
     ["pre4at","4AT recorded on admission"],
     ["geriatric","Senior geriatrician assessment within 72 hours"],
     ["physio","Physiotherapy assessment by day after surgery"],
     ["mobilised","Mobilised out of bed by day after surgery"],
     ["nutrition","Nutritional risk assessment"],
     ["delirium","Not delirious when tested post-op"],
     ["falls","Falls assessment completed"],
     ["bone","Bone health assessment/action"],
     ["bpt","NHFD best-practice criteria met"]
   ];
   const metricResults=metrics.map(([key,label])=>({key,label,...nhfdPerformance(periodPatients,key)}));
   const average=metricResults.filter(m=>m.percentage!==null);
   const overall=average.length?Math.round(average.reduce((a,m)=>a+m.percentage,0)/average.length):null;
   const noDates=nof.length-periodPatients.length;
   const runMetric=(NHFD_RUN_METRICS.find(([key])=>key===nofRunMetric)||NHFD_RUN_METRICS[NHFD_RUN_METRICS.length-1]);
   const runData=nhfdMonthlyRunData(nof,nofRunMetric,12);
   return <div className="space-y-5">
   <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between"><SectionTitle title="NOF Dashboard" subtitle="Synthetic UAT cohort — 10 sample NOF patients with populated NHFD v16 data for performance demonstration"/><div className="flex flex-wrap gap-2"><label className="flex items-center gap-2 rounded-xl border bg-white px-3 py-2 text-sm font-semibold"><span className="text-slate-500">Reporting period</span><select value={nofPerformancePeriod} onChange={e=>setNofPerformancePeriod(e.target.value)} className="rounded-lg border-0 bg-slate-50 px-2 py-1 font-bold text-violet-800 outline-none"><option value="24h">Last 24 hours</option><option value="7d">Last 1 week</option><option value="30d">Last 30 days</option><option value="12m">Last 12 months</option></select></label><Button onClick={openAddNofPatient} className="bg-violet-700 hover:bg-violet-800"><Plus className="mr-2 h-4 w-4"/>Add NOF patient</Button><Badge className="bg-violet-100 px-3 py-2 text-violet-800">{nof.length} sample NOF cases · {nofSelectedIds.length} selected</Badge><Button variant="outline" onClick={selectAllNof}><CheckSquare className="mr-2 h-4 w-4"/>Select all NOF</Button><Button variant="outline" onClick={clearNofSelection}>Clear selection</Button><Button disabled={!nofSelectedIds.length} onClick={exportNHFD} className="bg-emerald-700 hover:bg-emerald-800"><FileSpreadsheet className="mr-2 h-4 w-4"/>Export selected for NHFD</Button></div></div>
   <Card className="rounded-2xl border-violet-200 bg-violet-50/60"><CardContent className="p-4"><div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between"><div><b className="text-violet-950">NHFD performance — {periodLabels[nofPerformancePeriod]}</b><p className="mt-1 text-sm text-violet-900/80">Percentages are calculated from NOF records whose NHFD presentation date falls within the selected period.</p></div><div className="rounded-xl bg-white px-4 py-3 text-center shadow-sm"><div className="text-xs font-bold uppercase tracking-wide text-slate-500">Average across available measures</div><div className="text-2xl font-extrabold text-violet-800">{overall===null?"—":`${overall}%`}</div><div className="text-xs text-slate-500">{periodPatients.length} dated NOF case{periodPatients.length===1?"":"s"}</div></div></div>{noDates>0&&<div className="mt-3 rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900"><b>{noDates} NOF record{noDates===1?" has":"s have"} no NHFD presentation date.</b> Those records are excluded from the selected-period percentages. Complete the NHFD admission fields to make them appear in time-period reporting.</div>}</CardContent></Card>
   <div className="grid grid-cols-2 gap-3 md:grid-cols-3 xl:grid-cols-6"><KPI icon={HeartPulse} label="NOF cohort" value={nof.length} tone="violet"/><KPI icon={AlertTriangle} label="4AT ≥4" value={nof.filter(p=>p.fourAT>=4).length} tone="red"/><KPI icon={UserRound} label="OG outstanding" value={nof.filter(p=>!p.og).length} tone="amber"/><KPI icon={Activity} label="Bone health due" value={nof.filter(p=>!p.bone).length} tone="amber"/><KPI icon={ClipboardCheck} label="Falls due" value={nof.filter(p=>!p.falls).length} tone="amber"/><KPI icon={CheckCircle2} label="NHFD ready" value={nof.filter(p=>p.nhfd==="Ready").length} tone="green"/></div>
   <Card className="rounded-2xl"><CardHeader><CardTitle className="flex items-center justify-between text-base"><span>NHFD best-practice performance</span><Badge variant="outline">{periodLabels[nofPerformancePeriod]}</Badge></CardTitle><p className="text-sm text-slate-500">Percentage of eligible cases meeting each locally calculated NHFD performance measure. Eligibility follows the published NHFD measure definitions where the required fields are available.</p></CardHeader><CardContent className="space-y-4">{metricResults.map(m=><div key={m.key} className="rounded-xl border bg-white p-3"><div className="flex items-center justify-between gap-3"><div className="min-w-0"><div className="font-semibold text-slate-800">{m.label}</div><div className="text-xs text-slate-500">{m.percentage===null?"Insufficient dated/eligible data":`${m.passed} of ${m.eligible} eligible cases`}</div></div><div className="shrink-0 text-right"><div className="text-lg font-extrabold text-violet-800">{m.percentage===null?"—":`${m.percentage}%`}</div></div></div><div className="mt-2 h-3 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full bg-violet-600 transition-all" style={{width:`${m.percentage===null?0:m.percentage}%`}}/></div></div>)}</CardContent></Card>
   <Card className="rounded-2xl border-slate-200">
    <CardHeader>
      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <div><CardTitle className="text-base">Monthly NHFD-style run chart</CardTitle><p className="mt-1 text-sm text-slate-500">Monthly percentage performance using the Trust discharge month as the cohort time-base, mirroring the NHFD best-practice chart methodology.</p></div>
        <label className="flex items-center gap-2 rounded-xl border bg-slate-50 px-3 py-2 text-sm font-semibold"><span className="text-slate-500">Measure</span><select value={nofRunMetric} onChange={e=>setNofRunMetric(e.target.value)} className="rounded-lg border bg-white px-2 py-1 font-bold text-violet-800 outline-none">{NHFD_RUN_METRICS.map(([key,label])=><option key={key} value={key}>{label}</option>)}</select></label>
      </div>
    </CardHeader>
    <CardContent>
      <div className="overflow-x-auto">
        <div className="min-w-[820px]">
          <div className="relative h-72 rounded-xl border bg-white p-4">
            {[0,25,50,75,100].map(v=><div key={v} className="absolute left-12 right-4 border-t border-dashed border-slate-200" style={{top:`${16+(100-v)*0.64}%`}}><span className="absolute -left-10 -top-2 w-8 text-right text-[10px] text-slate-400">{v}%</span></div>)}
            <svg viewBox="0 0 1000 250" preserveAspectRatio="none" className="absolute inset-x-12 top-4 h-56 w-[calc(100%-3rem)] overflow-visible">
              {runData.length>1 && <polyline fill="none" stroke={NHFD_RUN_COLORS[NHFD_RUN_METRICS.findIndex(([key])=>key===nofRunMetric)]||"#111827"} strokeWidth="3" strokeLinejoin="round" strokeLinecap="round" points={runData.map((d,i)=>{const x=i*(1000/(runData.length-1));const y=d.percentage===null?245:245-(d.percentage*2.25);return `${x},${y}`;}).join(" ")}/> }
              {runData.map((d,i)=>{const x=runData.length===1?500:i*(1000/(runData.length-1));const y=d.percentage===null?245:245-(d.percentage*2.25);const color=NHFD_RUN_COLORS[NHFD_RUN_METRICS.findIndex(([key])=>key===nofRunMetric)]||"#111827";return <g key={`${d.year}-${d.month}`}><circle cx={x} cy={y} r="5" fill={d.percentage===null?"#cbd5e1":color}/>{d.percentage!==null&&<text x={x} y={Math.max(14,y-10)} textAnchor="middle" fontSize="11" fontWeight="700" fill={color}>{d.percentage}%</text>}</g>})}
            </svg>
            <div className="absolute bottom-2 left-12 right-4 flex justify-between text-[10px] font-semibold text-slate-500">{runData.map(d=><span key={`${d.year}-${d.month}`} className="min-w-0 text-center">{d.label}</span>)}</div>
          </div>
          <div className="mt-3 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500"><span><b className="text-slate-700">Selected measure:</b> {runMetric[1]}</span><span>Hover-ready data are also shown below as monthly cases and percentages.</span></div>
          <div className="mt-3 overflow-x-auto rounded-xl border">
            <table className="w-full min-w-[760px] text-xs"><thead className="bg-slate-50"><tr><th className="px-3 py-2 text-left">Month</th><th className="px-3 py-2 text-right">Eligible</th><th className="px-3 py-2 text-right">Achieved</th><th className="px-3 py-2 text-right">Performance</th></tr></thead><tbody className="divide-y">{runData.map(d=><tr key={`${d.year}-${d.month}`}><td className="px-3 py-2 font-semibold">{d.label}</td><td className="px-3 py-2 text-right">{d.eligible||0}</td><td className="px-3 py-2 text-right">{d.passed||0}</td><td className="px-3 py-2 text-right font-bold">{d.percentage===null?"—":`${d.percentage}%`}</td></tr>)}</tbody></table>
          </div>
          <div className="mt-3 rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900"><b>Methodology:</b> NHFD's published Best Practice run chart uses Trust Discharge Date as its monthly time-base and excludes records with errors, duplicates, conflicts and other specified exclusions. This local dashboard uses the same discharge-month principle, but its calculations remain operational/local and are not the official NHFD result.</div>
        </div>
      </div>
    </CardContent>
   </Card>
   <Card className="rounded-2xl border-slate-200"><CardHeader><CardTitle className="text-base">NHFD measures and interpretation</CardTitle></CardHeader><CardContent className="grid gap-3 md:grid-cols-2"><div className="rounded-xl bg-slate-50 p-4 text-sm"><b>Assessment</b><p className="mt-1 text-slate-600">Ward admission within 4 hours, admission 4AT, senior geriatric review within 72 hours, nutrition, falls and bone health.</p></div><div className="rounded-xl bg-slate-50 p-4 text-sm"><b>Post-operative care</b><p className="mt-1 text-slate-600">Physiotherapy, mobilisation by the day after surgery and post-operative delirium assessment.</p></div><div className="rounded-xl bg-slate-50 p-4 text-sm"><b>Best-practice tariff</b><p className="mt-1 text-slate-600">The dashboard provides a local calculation using the NHFD BPT v3 component fields recorded in this application.</p></div><div className="rounded-xl bg-amber-50 p-4 text-sm"><b>Important</b><p className="mt-1 text-amber-900">This is a local operational dashboard. NHFD’s own calculations, exclusions and follow-up timing remain authoritative for formal reporting and submission.</p></div></CardContent></Card>
   <Card className="overflow-hidden rounded-2xl"><CardContent className="p-0"><div className="overflow-x-auto"><table className="w-full min-w-[1120px] text-left text-sm"><thead className="bg-violet-950 text-xs uppercase text-white"><tr><th className="w-12 px-4 py-3">Select</th><th className="px-4 py-3">Patient</th><th className="px-4 py-3">Location</th><th className="px-4 py-3">Diagnosis</th><th className="px-4 py-3">4AT</th><th className="px-4 py-3">Orthogeriatrics</th><th className="px-4 py-3">Bone health</th><th className="px-4 py-3">Falls</th><th className="px-4 py-3">NHFD</th><th className="px-4 py-3">NHFD data</th></tr></thead><tbody className="divide-y">{nof.map(p=>{const checked=nofSelectedIds.includes(p.id);return <tr key={p.id} className={`hover:bg-violet-50 ${checked?"bg-violet-50/70":""}`}><td className="px-4 py-3"><button aria-label={checked?`Deselect ${p.name}`:`Select ${p.name}`} onClick={()=>toggleNofSelection(p.id)} className="rounded-md p-1 text-violet-700 hover:bg-violet-100">{checked?<CheckSquare className="h-5 w-5"/>:<Square className="h-5 w-5"/>}</button></td><td className="px-4 py-3"><button onClick={()=>openPatient(p)} className="text-left"><b>{p.name}</b><div className="text-xs text-slate-500">{p.nhs}</div></button></td><td className="px-4 py-3">{p.ward} · {p.bed}</td><td className="px-4 py-3">{p.diagnosis}</td><td className="px-4 py-3"><Badge className={p.fourAT>=4?"bg-red-100 text-red-800":"bg-emerald-100 text-emerald-800"}>{p.fourAT}</Badge></td><td className="px-4 py-3">{p.og?"✓ Complete":"⚠ Outstanding"}</td><td className="px-4 py-3">{p.bone?"✓ Complete":"⚠ Outstanding"}</td><td className="px-4 py-3">{p.falls?"✓ Complete":"⚠ Outstanding"}</td><td className="px-4 py-3"><Badge variant="outline">{p.nhfd}</Badge></td><td className="px-4 py-3"><Button variant="outline" onClick={()=>openEditNofPatient(p)}><Pencil className="mr-1 h-4 w-4"/>Edit</Button></td></tr>})}</tbody></table></div></CardContent></Card>
 </div>;
 };

 const TheatreBoard=()=> <div className="space-y-5"><SectionTitle title="Theatre Planning Board" subtitle="Synthetic list order, procedures and readiness overview"/><div className="grid gap-5 xl:grid-cols-3">{["Theatre 1","Theatre 2","Backup"].map(t=><Card key={t} className="rounded-2xl"><CardHeader><CardTitle className="flex items-center justify-between text-base"><span>{t}</span><Badge>{patientRecords.filter(p=>p.theatre===t).length} cases</Badge></CardTitle></CardHeader><CardContent className="space-y-3">{patientRecords.filter(p=>p.theatre===t).map((p,i)=><button key={p.id} onClick={()=>openPatient(p)} className="w-full rounded-xl border p-4 text-left hover:border-blue-300 hover:shadow-md"><div className="flex items-center gap-3"><span className="flex h-9 w-9 items-center justify-center rounded-xl bg-[#005eb8] font-bold text-white">{i+1}</span><div className="min-w-0 flex-1"><b className="block truncate">{p.name}</b><span className="block truncate text-xs text-slate-500">{p.diagnosis}</span></div><Badge className={`${pClass[p.priority]} border`}>{p.priority}</Badge></div><div className="mt-3 rounded-lg bg-slate-50 p-2 text-xs"><b>{p.procedure}</b><br/>Status: {p.status}</div></button>)}</CardContent></Card>)}</div></div>;

 const MDT=()=> <div className="space-y-5"><SectionTitle title="MDT / Handover Workspace" subtitle={`${selectedForHandover.length} selected patient${selectedForHandover.length===1?"":"s"} for handover review`}/><div className="grid gap-5 xl:grid-cols-[1fr_360px]"><Card className="rounded-2xl"><CardHeader><CardTitle className="text-base">Patients to discuss</CardTitle><p className="text-sm text-slate-500">Only patients selected from the Trauma Patient List are shown here.</p></CardHeader><CardContent className="grid gap-2 md:grid-cols-2">{selectedForHandover.length ? selectedForHandover.map(p=><button key={p.id} onClick={()=>openPatient(p)} className="flex items-center justify-between rounded-xl border p-3 text-left hover:bg-blue-50"><div><b className="text-sm">{p.name}</b><div className="text-xs text-slate-500">{p.ward} · {p.diagnosis}</div></div><Badge className={`${pClass[p.priority]} border`}>{p.priority}</Badge></button>) : <div className="md:col-span-2 rounded-xl border border-dashed p-8 text-center text-sm text-slate-500">No patients are selected for handover. Return to the Trauma Patient List and tick the patients you want to discuss.</div>}</CardContent></Card><Card className="h-fit rounded-2xl"><CardHeader><CardTitle className="text-base">Actions due</CardTitle></CardHeader><CardContent className="space-y-3">{patientRecords.slice().sort((a,b)=>a.due.localeCompare(b.due)).slice(0,8).map(p=><button key={p.id} onClick={()=>openPatient(p)} className="w-full rounded-xl bg-slate-50 p-3 text-left hover:bg-blue-50"><div className="flex justify-between"><b className="text-sm">{p.name}</b><span className="text-xs font-bold text-red-600">{p.due}</span></div><div className="mt-1 text-xs text-slate-600">{p.owner}: {p.action.split(';')[0]}</div></button>)}</CardContent></Card></div></div>;

 const content={"Command Centre":CommandCentre(),"Patients Admission List":TraumaList(),"Patient Detail":PatientDetail(),"NOF Dashboard":NofDashboard(),"Theatre Board":TheatreBoard(),"MDT / Handover":MDT()}[screen];
 return <div className="flex min-h-screen bg-slate-100 text-slate-950">{QuickActionModal()}{PatientEditor()}{NofPatientForm()}{Sidebar()}{mobileOpen&&<button className="fixed inset-0 z-30 bg-black/40 lg:hidden" onClick={()=>setMobileOpen(false)}/>}<div className="min-w-0 flex-1"><header className="sticky top-0 z-20 flex h-20 items-center justify-between border-b bg-white/95 px-4 backdrop-blur md:px-6"><div className="flex items-center gap-3"><button className="rounded-xl border p-2 lg:hidden" onClick={()=>setMobileOpen(true)}><Menu/></button><div><div className="text-xs font-bold uppercase tracking-widest text-[#005eb8]">Worcestershire Acute Hospitals NHS Trust</div><div className="font-bold">Trauma & Orthopaedics Digital Handover</div></div></div><div className="hidden items-center gap-2 md:flex"><Badge variant="outline"><MapPin className="mr-1 h-3 w-3"/>WRH</Badge><Badge variant="outline"><Clock3 className="mr-1 h-3 w-3"/>UAT session</Badge></div></header><main className="p-4 md:p-6"> <AnimatePresence mode="wait"><motion.div key={screen} initial={{opacity:0,y:8}} animate={{opacity:1,y:0}} exit={{opacity:0,y:-5}} transition={{duration:.18}}>{content}</motion.div></AnimatePresence></main></div></div>;
}