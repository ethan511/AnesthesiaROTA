import { useState, useEffect, useMemo, useRef, createContext, useContext } from "react";
import {
  Sun, Moon, Clock, Phone, X, Check, CalendarDays, List as ListIcon, StickyNote,
  Upload, Printer, CalendarPlus, Loader2, AlertCircle,
} from "lucide-react";
import * as XLSX from "xlsx";
 
/* ---------- Default / seed data (whatever was last loaded, e.g. current month) ---------- */
const DEFAULT_ROTA = {"monthLabel":"July 2026","year":2026,"monthIndex":6,"days":[{"date":1,"day":"WED","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"M. Zahrani"},{"role":"L&D Shift (Night)","name":"Naif"},{"role":"Trauma Shift (Day)","name":"Qahtani"},{"role":"Trauma Shift (Night)","name":"Bahaa"},{"role":"Main OR","name":"Sami Z"},{"role":"Main OR","name":"Sami A"},{"role":"DSU","name":"Hatim"},{"role":"DSU","name":"Mulham"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"Raoum"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sondos"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Dosari"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Rayan"},{"role":"Sedation/Radiology & Endo","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Ashraf"},{"role":"Business Coverage","name":"Sondos"}]},{"date":2,"day":"THU","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"M. Zahrani"},{"role":"L&D Shift (Night)","name":"Naif"},{"role":"Trauma Shift (Day)","name":"Qahtani"},{"role":"Trauma Shift (Night)","name":"Bahaa"},{"role":"Main OR","name":"Sami Z"},{"role":"Main OR","name":"Aymen"},{"role":"DSU","name":"Hatim"},{"role":"DSU","name":"Mulham"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"Raoum"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sondos"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Dosari"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Rayan"},{"role":"Sedation/Radiology & Endo","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Ashraf"},{"role":"Business Coverage","name":"Sondos"}]},{"date":3,"day":"FRI","assignments":[{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"M. Zahrani"},{"role":"L&D Shift (Night)","name":"Naif"},{"role":"Trauma Shift (Day)","name":"Hani"},{"role":"Trauma Shift (Night)","name":"Yonis"}]},{"date":4,"day":"SAT","assignments":[{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"M. Zahrani"},{"role":"L&D Shift (Night)","name":"Naif"},{"role":"Trauma Shift (Day)","name":"Khaled"},{"role":"Trauma Shift (Night)","name":"Hani"},{"role":"Saturday List","name":"Hakami"},{"role":"Saturday List","name":"Dosari"},{"role":"Saturday List","name":"Al-Osaimi"},{"role":"Saturday List","name":"Salmi"},{"role":"Business Coverage","name":"Mulham"}]},{"date":5,"day":"SUN","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Khaled"},{"role":"Trauma Shift (Night)","name":"Hani"},{"role":"Main OR","name":"Roqi"},{"role":"Main OR","name":"Mulham"},{"role":"DSU","name":"Yonis"},{"role":"DSU","name":"Sondos"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Aymen"},{"role":"Sedation/Radiology & Endo","name":"Qahtani"},{"role":"Sedation/Radiology & Endo","name":"Hatim"},{"role":"Business Coverage","name":"Mulham"}]},{"date":6,"day":"MON","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Khaled"},{"role":"Trauma Shift (Night)","name":"Hani"},{"role":"Main OR","name":"Roqi"},{"role":"Main OR","name":"Mulham"},{"role":"DSU","name":"Yonis"},{"role":"DSU","name":"Sondos"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Aymen"},{"role":"Sedation/Radiology & Endo","name":"Qahtani"},{"role":"Sedation/Radiology & Endo","name":"Hatim"},{"role":"Business Coverage","name":"Mulham"}]},{"date":7,"day":"TUE","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Khaled"},{"role":"Trauma Shift (Night)","name":"Hani"},{"role":"Main OR","name":"Roqi"},{"role":"Main OR","name":"Mulham"},{"role":"DSU","name":"Yonis"},{"role":"DSU","name":"Sondos"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Aymen"},{"role":"Sedation/Radiology & Endo","name":"Qahtani"},{"role":"Sedation/Radiology & Endo","name":"Hatim"},{"role":"Business Coverage","name":"Mulham"}]},{"date":8,"day":"WED","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Khaled"},{"role":"Trauma Shift (Night)","name":"Hani"},{"role":"Main OR","name":"Roqi"},{"role":"Main OR","name":"Mulham"},{"role":"DSU","name":"Yonis"},{"role":"DSU","name":"Sondos"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Aymen"},{"role":"Sedation/Radiology & Endo","name":"Qahtani"},{"role":"Sedation/Radiology & Endo","name":"Hatim"},{"role":"Business Coverage","name":"Mulham"}]},{"date":9,"day":"THU","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Khaled"},{"role":"Trauma Shift (Night)","name":"Hani"},{"role":"Main OR","name":"Roqi"},{"role":"Main OR","name":"Mulham"},{"role":"DSU","name":"Yonis"},{"role":"DSU","name":"Sondos"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Firash"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Ahmed"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Aymen"},{"role":"Sedation/Radiology & Endo","name":"Qahtani"},{"role":"Sedation/Radiology & Endo","name":"Hatim"},{"role":"Business Coverage","name":"Mulham"}]},{"date":10,"day":"FRI","assignments":[{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Essa"},{"role":"Trauma Shift (Night)","name":"Sondos"}]},{"date":11,"day":"SAT","assignments":[{"role":"Tech In-Charge","name":"Sami A"},{"role":"L&D Shift (Day)","name":"Raoum"},{"role":"L&D Shift (Night)","name":"Dosari"},{"role":"Trauma Shift (Day)","name":"Moh"},{"role":"Trauma Shift (Night)","name":"Ahmed"},{"role":"Saturday List","name":"Roqi"},{"role":"Saturday List","name":"Aymen"},{"role":"Saturday List","name":"Bahaa"},{"role":"Saturday List","name":"Qahtani"},{"role":"Business Coverage","name":"Al-Osaimi"}]},{"date":12,"day":"SUN","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Moh"},{"role":"Trauma Shift (Night)","name":"Ahmed"},{"role":"Main OR","name":"Bahaa"},{"role":"Main OR","name":"Naif"},{"role":"DSU","name":"Firash"},{"role":"DSU","name":"Al-Osaimi"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"M. Zahrani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Yonis"},{"role":"Sedation/Radiology & Endo","name":"Sondos"},{"role":"Sedation/Radiology & Endo","name":"Fawzi"},{"role":"Business Coverage","name":"Al-Osaimi"}]},{"date":13,"day":"MON","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Moh"},{"role":"Trauma Shift (Night)","name":"Ahmed"},{"role":"Main OR","name":"Bahaa"},{"role":"Main OR","name":"Naif"},{"role":"DSU","name":"Firash"},{"role":"DSU","name":"Al-Osaimi"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"M. Zahrani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Yonis"},{"role":"Sedation/Radiology & Endo","name":"Sondos"},{"role":"Sedation/Radiology & Endo","name":"Fawzi"},{"role":"Business Coverage","name":"Al-Osaimi"}]},{"date":14,"day":"TUE","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Moh"},{"role":"Trauma Shift (Night)","name":"Ahmed"},{"role":"Main OR","name":"Bahaa"},{"role":"Main OR","name":"Naif"},{"role":"DSU","name":"Firash"},{"role":"DSU","name":"Al-Osaimi"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"M. Zahrani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Yonis"},{"role":"Sedation/Radiology & Endo","name":"Sondos"},{"role":"Sedation/Radiology & Endo","name":"Fawzi"},{"role":"Business Coverage","name":"Al-Osaimi"}]},{"date":15,"day":"WED","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Moh"},{"role":"Trauma Shift (Night)","name":"Ahmed"},{"role":"Main OR","name":"Bahaa"},{"role":"Main OR","name":"Naif"},{"role":"DSU","name":"Firash"},{"role":"DSU","name":"Al-Osaimi"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"M. Zahrani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Yonis"},{"role":"Sedation/Radiology & Endo","name":"Sondos"},{"role":"Sedation/Radiology & Endo","name":"Fawzi"},{"role":"Business Coverage","name":"Al-Osaimi"}]},{"date":16,"day":"THU","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Moh"},{"role":"Trauma Shift (Night)","name":"Ahmed"},{"role":"Main OR","name":"Bahaa"},{"role":"Main OR","name":"Naif"},{"role":"DSU","name":"Firash"},{"role":"DSU","name":"Al-Osaimi"},{"role":"Burn Unit","name":"Adel"},{"role":"NSTC","name":"M. Zahrani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Salmi"},{"role":"NSTC","name":"Yonis"},{"role":"Sedation/Radiology & Endo","name":"Sondos"},{"role":"Sedation/Radiology & Endo","name":"Fawzi"},{"role":"Business Coverage","name":"Al-Osaimi"}]},{"date":17,"day":"FRI","assignments":[{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Bahaa"},{"role":"Trauma Shift (Night)","name":"Hatim"}]},{"date":18,"day":"SAT","assignments":[{"role":"Tech In-Charge","name":"Sami Z"},{"role":"L&D Shift (Day)","name":"Aymen"},{"role":"L&D Shift (Night)","name":"Qahtani"},{"role":"Trauma Shift (Day)","name":"Al-Osaimi"},{"role":"Trauma Shift (Night)","name":"Essa"},{"role":"Saturday List","name":"Firash"},{"role":"Saturday List","name":"Saqa"},{"role":"Saturday List","name":"Fawzi"},{"role":"Saturday List","name":"Hatim"},{"role":"Business Coverage","name":"M. Zahrani"}]},{"date":19,"day":"SUN","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Al-Osaimi"},{"role":"Trauma Shift (Night)","name":"Essa"},{"role":"Main OR","name":"Hatim"},{"role":"Main OR","name":"Sondos"},{"role":"DSU","name":"Rana"},{"role":"DSU","name":"M. Zahrani"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Bahaa"},{"role":"Sedation/Radiology & Endo","name":"Ahmed"},{"role":"Sedation/Radiology & Endo","name":"Firash"},{"role":"Business Coverage","name":"M. Zahrani"}]},{"date":20,"day":"MON","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Al-Osaimi"},{"role":"Trauma Shift (Night)","name":"Essa"},{"role":"Main OR","name":"Hatim"},{"role":"Main OR","name":"Sondos"},{"role":"DSU","name":"Rana"},{"role":"DSU","name":"M. Zahrani"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Bahaa"},{"role":"Sedation/Radiology & Endo","name":"Ahmed"},{"role":"Sedation/Radiology & Endo","name":"Firash"},{"role":"Business Coverage","name":"M. Zahrani"}]},{"date":21,"day":"TUE","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Al-Osaimi"},{"role":"Trauma Shift (Night)","name":"Essa"},{"role":"Main OR","name":"Hatim"},{"role":"Main OR","name":"Sondos"},{"role":"DSU","name":"Rana"},{"role":"DSU","name":"M. Zahrani"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Bahaa"},{"role":"Sedation/Radiology & Endo","name":"Ahmed"},{"role":"Sedation/Radiology & Endo","name":"Firash"},{"role":"Business Coverage","name":"M. Zahrani"}]},{"date":22,"day":"WED","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Al-Osaimi"},{"role":"Trauma Shift (Night)","name":"Essa"},{"role":"Main OR","name":"Hatim"},{"role":"Main OR","name":"Sondos"},{"role":"DSU","name":"Rana"},{"role":"DSU","name":"M. Zahrani"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Bahaa"},{"role":"Sedation/Radiology & Endo","name":"Ahmed"},{"role":"Sedation/Radiology & Endo","name":"Firash"},{"role":"Business Coverage","name":"M. Zahrani"}]},{"date":23,"day":"THU","assignments":[{"role":"In Charge","name":"Sami Z"},{"role":"Main OR Coordinator","name":"Sami Z"},{"role":"Outside Coordinator","name":"Sami Z"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Sameer"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Al-Osaimi"},{"role":"Trauma Shift (Night)","name":"Essa"},{"role":"Main OR","name":"Hatim"},{"role":"Main OR","name":"Sondos"},{"role":"DSU","name":"Rana"},{"role":"DSU","name":"M. Zahrani"},{"role":"Burn Unit","name":"Sami A"},{"role":"NSTC","name":"Roqi"},{"role":"NSTC","name":"Adel"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Fawzi"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Sameer"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Khaled"},{"role":"NSTC","name":"Bahaa"},{"role":"Sedation/Radiology & Endo","name":"Ahmed"},{"role":"Sedation/Radiology & Endo","name":"Firash"},{"role":"Business Coverage","name":"M. Zahrani"}]},{"date":24,"day":"FRI","assignments":[{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Hakami"},{"role":"Trauma Shift (Night)","name":"Al-Osaimi"}]},{"date":25,"day":"SAT","assignments":[{"role":"Tech In-Charge","name":"Sameer"},{"role":"L&D Shift (Day)","name":"Dosari"},{"role":"L&D Shift (Night)","name":"Raoum"},{"role":"Trauma Shift (Day)","name":"Roqi"},{"role":"Trauma Shift (Night)","name":"Adel"},{"role":"Saturday List","name":"Rana"},{"role":"Saturday List","name":"Sondos"},{"role":"Saturday List","name":"M. Zahrani"},{"role":"Saturday List","name":"Essa"},{"role":"Business Coverage","name":"Ahmed"}]},{"date":26,"day":"SUN","assignments":[{"role":"In Charge","name":"Saeed"},{"role":"Main OR Coordinator","name":"Sami A"},{"role":"Outside Coordinator","name":"Sami A"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Rayan"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Rayan"},{"role":"L&D Shift (Day)","name":"Qahtani"},{"role":"L&D Shift (Night)","name":"Aymen"},{"role":"Trauma Shift (Day)","name":"Roqi"},{"role":"Trauma Shift (Night)","name":"Adel"},{"role":"Main OR","name":"Fawzi"},{"role":"Main OR","name":"Firash"},{"role":"DSU","name":"Sondos"},{"role":"DSU","name":"Ahmed"},{"role":"Burn Unit","name":"Saeed"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Rana"},{"role":"Sedation/Radiology & Endo","name":"M. Zahrani"},{"role":"Business Coverage","name":"Ahmed"}]},{"date":27,"day":"MON","assignments":[{"role":"In Charge","name":"Saeed"},{"role":"Main OR Coordinator","name":"Sami A"},{"role":"Outside Coordinator","name":"Sami A"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Rayan"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Rayan"},{"role":"L&D Shift (Day)","name":"Qahtani"},{"role":"L&D Shift (Night)","name":"Aymen"},{"role":"Trauma Shift (Day)","name":"Roqi"},{"role":"Trauma Shift (Night)","name":"Adel"},{"role":"Main OR","name":"Fawzi"},{"role":"Main OR","name":"Firash"},{"role":"DSU","name":"Sondos"},{"role":"DSU","name":"Ahmed"},{"role":"Burn Unit","name":"Saeed"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Rana"},{"role":"Sedation/Radiology & Endo","name":"M. Zahrani"},{"role":"Business Coverage","name":"Ahmed"}]},{"date":28,"day":"TUE","assignments":[{"role":"In Charge","name":"Saeed"},{"role":"Main OR Coordinator","name":"Sami A"},{"role":"Outside Coordinator","name":"Sami A"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Rayan"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Rayan"},{"role":"L&D Shift (Day)","name":"Qahtani"},{"role":"L&D Shift (Night)","name":"Aymen"},{"role":"Trauma Shift (Day)","name":"Roqi"},{"role":"Trauma Shift (Night)","name":"Adel"},{"role":"Main OR","name":"Fawzi"},{"role":"Main OR","name":"Firash"},{"role":"DSU","name":"Sondos"},{"role":"DSU","name":"Ahmed"},{"role":"Burn Unit","name":"Saeed"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Rana"},{"role":"Sedation/Radiology & Endo","name":"M. Zahrani"},{"role":"Business Coverage","name":"Ahmed"}]},{"date":29,"day":"WED","assignments":[{"role":"In Charge","name":"Saeed"},{"role":"Main OR Coordinator","name":"Sami A"},{"role":"Outside Coordinator","name":"Sami A"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Rayan"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Rayan"},{"role":"L&D Shift (Day)","name":"Qahtani"},{"role":"L&D Shift (Night)","name":"Aymen"},{"role":"Trauma Shift (Day)","name":"Roqi"},{"role":"Trauma Shift (Night)","name":"Adel"},{"role":"Main OR","name":"Fawzi"},{"role":"Main OR","name":"Firash"},{"role":"DSU","name":"Sondos"},{"role":"DSU","name":"Ahmed"},{"role":"Burn Unit","name":"Saeed"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Rana"},{"role":"Sedation/Radiology & Endo","name":"M. Zahrani"},{"role":"Business Coverage","name":"Ahmed"}]},{"date":30,"day":"THU","assignments":[{"role":"In Charge","name":"Saeed"},{"role":"Main OR Coordinator","name":"Sami A"},{"role":"Outside Coordinator","name":"Sami A"},{"role":"SCDDP Coordinator","name":"Rayan"},{"role":"NSTC Coordinator","name":"Rayan"},{"role":"Supply Coordinator","name":"Sami A"},{"role":"Tech In-Charge","name":"Rayan"},{"role":"L&D Shift (Day)","name":"Qahtani"},{"role":"L&D Shift (Night)","name":"Aymen"},{"role":"Trauma Shift (Day)","name":"Roqi"},{"role":"Trauma Shift (Night)","name":"Adel"},{"role":"Main OR","name":"Fawzi"},{"role":"Main OR","name":"Firash"},{"role":"DSU","name":"Sondos"},{"role":"DSU","name":"Ahmed"},{"role":"Burn Unit","name":"Saeed"},{"role":"NSTC","name":"Al-Osaimi"},{"role":"NSTC","name":"Hani"},{"role":"NSTC","name":"Sami A"},{"role":"NSTC","name":"Essa"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Yonis"},{"role":"NSTC","name":"Rayan"},{"role":"NSTC","name":"Hatim"},{"role":"NSTC","name":"Mulham"},{"role":"NSTC","name":"Bahaa"},{"role":"NSTC","name":"Moh"},{"role":"NSTC","name":"Khaled"},{"role":"Sedation/Radiology & Endo","name":"Rana"},{"role":"Sedation/Radiology & Endo","name":"M. Zahrani"},{"role":"Business Coverage","name":"Ahmed"}]},{"date":31,"day":"FRI","assignments":[{"role":"Tech In-Charge","name":"Rayan"},{"role":"L&D Shift (Day)","name":"Qahtani"},{"role":"L&D Shift (Night)","name":"Aymen"},{"role":"Trauma Shift (Day)","name":"M. Zahrani"},{"role":"Trauma Shift (Night)","name":"Mulham"}]}],"names":["Adel","Ahmed","Al-Osaimi","Ashraf","Aymen","Bahaa","Dosari","Essa","Fawzi","Firash","Hakami","Hani","Hatim","Khaled","M. Zahrani","Moh","Mulham","Naif","Qahtani","Rana","Raoum","Rayan","Roqi","Saeed","Salmi","Sameer","Sami A","Sami Z","Saqa","Sondos","Yonis"]}
 
const DEFAULT_CONTACTS = [
  { label: "Chief, Anesthesia Tech Services", bleep: "1479" },
  { label: "Supervisor — Sami Z", bleep: "2090" },
  { label: "Supervisor — Sami A", bleep: "2535" },
  { label: "Supervisor — Waleed", bleep: "1962" },
  { label: "Supervisor — Sameer", bleep: "2523" },
  { label: "A/Supervisor — Rayan", bleep: "2792" },
];
 
/* ---------- Structural constants describing the standard rota template ---------- */
/* Columns 2..36 of the sheet, in order. Same 35-role template used every month. */
const ROLE_COLUMNS = [
  "In Charge", "Main OR Coordinator", "Outside Coordinator", "SCDDP Coordinator", "NSTC Coordinator",
  "Supply Coordinator", "Tech In-Charge",
  "L&D Shift (Day)", "L&D Shift (Night)", "Trauma Shift (Day)", "Trauma Shift (Night)",
  "Main OR", "Main OR", "DSU", "DSU", "Burn Unit",
  "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC", "NSTC",
  "Sedation/Radiology & Endo", "Sedation/Radiology & Endo",
  "Saturday List", "Saturday List", "Saturday List", "Saturday List",
  "Business Coverage",
];
 
const ROLE_META = {
  "In Charge": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "Main OR Coordinator": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "Outside Coordinator": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "SCDDP Coordinator": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "NSTC Coordinator": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "Supply Coordinator": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "Tech In-Charge": { time: "07:00\u201316:00", shift: "day", category: "coordinator" },
  "L&D Shift (Day)": { time: "07:00\u201319:00", shift: "day", category: "ld" },
  "L&D Shift (Night)": { time: "19:00\u201307:00", shift: "night", category: "ld" },
  "Trauma Shift (Day)": { time: "07:00\u201319:00", shift: "day", category: "trauma" },
  "Trauma Shift (Night)": { time: "19:00\u201307:00", shift: "night", category: "trauma" },
  "Main OR": { time: "07:00\u201316:00", shift: "day", category: "theatre" },
  "DSU": { time: "07:00\u201316:00", shift: "day", category: "theatre" },
  "Burn Unit": { time: "07:00\u201316:00", shift: "day", category: "theatre" },
  "NSTC": { time: "07:00\u201316:00", shift: "day", category: "theatre" },
  "Sedation/Radiology & Endo": { time: "07:00\u201316:00", shift: "day", category: "theatre" },
  "Saturday List": { time: "07:00\u201316:00", shift: "day", category: "saturday" },
  "Business Coverage": { time: "24 hours", shift: "24h", category: "business" },
};
 
const CATEGORY_STYLE = {
  business: {
    dot: "bg-amber-500",
    light: { chipBg: "bg-amber-50", chipBorder: "border-amber-300", chipText: "text-amber-900", solid: "bg-amber-500" },
    dark: { chipBg: "bg-amber-950", chipBorder: "border-amber-700", chipText: "text-amber-300", solid: "bg-amber-500" },
  },
  saturday: {
    dot: "bg-sky-500",
    light: { chipBg: "bg-sky-50", chipBorder: "border-sky-300", chipText: "text-sky-900", solid: "bg-sky-500" },
    dark: { chipBg: "bg-sky-950", chipBorder: "border-sky-700", chipText: "text-sky-300", solid: "bg-sky-500" },
  },
  trauma: {
    dot: "bg-red-500",
    light: { chipBg: "bg-red-50", chipBorder: "border-red-300", chipText: "text-red-900", solid: "bg-red-500" },
    dark: { chipBg: "bg-red-950", chipBorder: "border-red-700", chipText: "text-red-300", solid: "bg-red-500" },
  },
  ld: {
    dot: "bg-rose-400",
    light: { chipBg: "bg-rose-50", chipBorder: "border-rose-300", chipText: "text-rose-900", solid: "bg-rose-400" },
    dark: { chipBg: "bg-rose-950", chipBorder: "border-rose-700", chipText: "text-rose-300", solid: "bg-rose-400" },
  },
  coordinator: {
    dot: "bg-violet-500",
    light: { chipBg: "bg-violet-50", chipBorder: "border-violet-300", chipText: "text-violet-900", solid: "bg-violet-500" },
    dark: { chipBg: "bg-violet-950", chipBorder: "border-violet-700", chipText: "text-violet-300", solid: "bg-violet-500" },
  },
  theatre: {
    dot: "bg-emerald-500",
    light: { chipBg: "bg-emerald-50", chipBorder: "border-emerald-300", chipText: "text-emerald-900", solid: "bg-emerald-500" },
    dark: { chipBg: "bg-emerald-950", chipBorder: "border-emerald-700", chipText: "text-emerald-300", solid: "bg-emerald-500" },
  },
};
const CATEGORY_ORDER = ["business", "saturday", "trauma", "ld", "coordinator", "theatre"];
const LEGEND_ITEMS = [
  ["business", "Business"], ["saturday", "Saturday list"], ["trauma", "Trauma"],
  ["ld", "L&D"], ["coordinator", "Coordinator"], ["theatre", "OR / DSU / NSTC"],
];
 
const DAY_COLS = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
const MONTH_ABBR = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
 
const FILTERS = [
  { key: "all", label: "All" },
  { key: "business", label: "Business" },
  { key: "saturday", label: "Saturday" },
  { key: "noted", label: "Noted" },
];
 
/* ---------- Date / key helpers ---------- */
function pad2(n) { return String(n).padStart(2, "0"); }
function dateKeyFor(year, monthIndex, date) {
  return year + "-" + pad2(monthIndex + 1) + "-" + pad2(date);
}
function fmtICSDate(year, monthIndex, date) {
  const d = new Date(year, monthIndex, date);
  return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate());
}
function nowStampUTC() {
  return new Date().toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
}
function escapeICS(s) {
  return String(s || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
}
 
/* ---------- Parsing an uploaded workbook into the same shape as DEFAULT_ROTA.days ---------- */
function parseGrid(grid) {
  const days = [];
  const namesSet = new Set();
  for (const row of grid) {
    if (!row || row.length < 2) continue;
    const rawDate = row[0];
    const dateNum = typeof rawDate === "number" ? rawDate : parseInt(rawDate, 10);
    const dayAbbr = String(row[1] == null ? "" : row[1]).trim().toUpperCase().slice(0, 3);
    if (!dateNum || dateNum < 1 || dateNum > 31) continue;
    if (DAY_COLS.indexOf(dayAbbr) === -1) continue;
    const assignments = [];
    for (let i = 0; i < ROLE_COLUMNS.length; i++) {
      const col = i + 2;
      const raw = row[col];
      const val = raw == null ? "" : String(raw).trim();
      if (val) {
        namesSet.add(val);
        assignments.push({ role: ROLE_COLUMNS[i], name: val });
      }
    }
    days.push({ date: dateNum, day: dayAbbr, assignments });
  }
  days.sort((a, b) => a.date - b.date);
  return { days, names: Array.from(namesSet).sort() };
}
 
function detectMonthYear(grid) {
  const monthRe = /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*-?\s*'?(\d{2,4})/i;
  for (const row of grid) {
    for (const cell of row) {
      if (typeof cell !== "string") continue;
      const m = cell.match(monthRe);
      if (m) {
        const idx = MONTH_ABBR.indexOf(m[1].toLowerCase().slice(0, 3));
        if (idx === -1) continue;
        let yr = parseInt(m[2], 10);
        if (yr < 100) yr += 2000;
        return yr + "-" + pad2(idx + 1);
      }
    }
  }
  return "";
}
 
/* ---------- PDF support (best-effort, client-side, no server involved) ----------
   Loads pdf.js from a CDN at runtime (not part of the bundled library set), reads
   text positions per page, clusters them into a row/column grid, then feeds that
   grid through the same parseGrid() used for Excel so both paths share one source
   of truth. This is inherently less certain than the Excel path: PDF text carries
   no real column metadata, only x/y coordinates, so column edges are inferred. */
let _pdfjsPromise = null;
function loadPdfJs() {
  if (typeof window !== "undefined" && window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (_pdfjsPromise) return _pdfjsPromise;
  const CDN_BASE = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/";
  _pdfjsPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = CDN_BASE + "pdf.min.js";
    script.onload = () => {
      if (window.pdfjsLib) {
        try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = CDN_BASE + "pdf.worker.min.js"; } catch (e) {}
        resolve(window.pdfjsLib);
      } else {
        reject(new Error("PDF reader script loaded but did not initialize."));
      }
    };
    script.onerror = () => reject(new Error("Could not load the PDF reader from the network."));
    document.head.appendChild(script);
  });
  const timeout = new Promise((_, reject) => window.setTimeout(() => reject(new Error("PDF reader took too long to load.")), 12000));
  return Promise.race([_pdfjsPromise, timeout]);
}
 
function clusterPositions(values, tolerance) {
  const sorted = [...values].sort((a, b) => a - b);
  const clusters = [];
  for (const v of sorted) {
    const last = clusters[clusters.length - 1];
    if (last && v - last.max <= tolerance) {
      last.max = v;
      last.sum += v;
      last.count += 1;
    } else {
      clusters.push({ max: v, sum: v, count: 1 });
    }
  }
  return clusters.map((c) => c.sum / c.count);
}
 
function nearestIndex(value, centers) {
  let bestIdx = 0, bestDist = Infinity;
  centers.forEach((c, i) => {
    const d = Math.abs(value - c);
    if (d < bestDist) { bestDist = d; bestIdx = i; }
  });
  return bestIdx;
}
 
const EXPECTED_COLS = ROLE_COLUMNS.length + 2; // date + day + 35 role columns
 
async function extractGridFromPdf(file) {
  const pdfjsLib = await loadPdfJs();
  const buf = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
 
  for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
    const page = await doc.getPage(pageNum);
    const content = await page.getTextContent();
    const items = content.items
      .map((it) => ({ text: (it.str || "").trim(), x: it.transform[4], y: it.transform[5] }))
      .filter((it) => it.text.length > 0);
    if (items.length === 0) continue;
 
    const byY = [...items].sort((a, b) => b.y - a.y);
    const rows = [];
    for (const it of byY) {
      let row = rows.find((r) => Math.abs(r.y - it.y) <= 3);
      if (!row) { row = { y: it.y, items: [] }; rows.push(row); }
      row.items.push(it);
    }
    rows.forEach((r) => r.items.sort((a, b) => a.x - b.x));
 
    const dayRows = rows.filter((r) => {
      if (r.items.length < 2) return false;
      const first = r.items[0].text;
      const second = r.items[1].text.toUpperCase().slice(0, 3);
      const n = parseInt(first, 10);
      return n >= 1 && n <= 31 && /^\d{1,2}$/.test(first) && DAY_COLS.indexOf(second) !== -1;
    });
    if (dayRows.length < 20) continue;
 
    const allX = [];
    dayRows.forEach((r) => r.items.forEach((it) => allX.push(it.x)));
    const centers = clusterPositions(allX, 6).sort((a, b) => a - b);
 
    if (centers.length !== EXPECTED_COLS) {
      throw new Error(
        "This PDF's columns didn't line up as expected (found " + centers.length + ", expected " + EXPECTED_COLS +
        "). PDF layout can shift slightly between exports — please use the .xlsx file instead, or send the PDF to Claude in chat."
      );
    }
 
    const grid = dayRows.map((r) => {
      const rowArr = new Array(centers.length).fill("");
      r.items.forEach((it) => {
        const idx = nearestIndex(it.x, centers);
        rowArr[idx] = rowArr[idx] ? rowArr[idx] + " " + it.text : it.text;
      });
      return rowArr;
    });
 
    return grid;
  }
 
  throw new Error("Couldn't find a day-by-day table on any page of this PDF.");
}
 
/* ---------- Export: .ics calendar file ---------- */
function generateICS(rota, name, dayMeta) {
  const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Anesthesia Technology Monthly Rota//EN", "CALSCALE:GREGORIAN"];
  rota.days.forEach((d) => {
    const mine = d.assignments.filter((a) => a.name === name);
    if (mine.length === 0) return;
    const start = fmtICSDate(rota.year, rota.monthIndex, d.date);
    const end = fmtICSDate(rota.year, rota.monthIndex, d.date + 1);
    const summary = mine.map((a) => a.role).join(", ");
    const meta = dayMeta[dateKeyFor(rota.year, rota.monthIndex, d.date)];
    let desc = mine.map((a) => a.role + " (" + ((ROLE_META[a.role] || {}).time || "") + ")").join("\\n");
    if (meta && meta.note) desc += "\\nNote: " + escapeICS(meta.note);
    lines.push("BEGIN:VEVENT");
    lines.push("UID:" + start + "-" + name.replace(/\s+/g, "") + "@anesthesia-rota");
    lines.push("DTSTAMP:" + nowStampUTC());
    lines.push("DTSTART;VALUE=DATE:" + start);
    lines.push("DTEND;VALUE=DATE:" + end);
    lines.push("SUMMARY:" + escapeICS(summary));
    lines.push("DESCRIPTION:" + desc);
    lines.push("END:VEVENT");
  });
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}
function downloadICS(rota, name, dayMeta) {
  const content = generateICS(rota, name, dayMeta);
  const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = (rota.monthLabel || "rota").replace(/\s+/g, "-") + "-" + name.replace(/\s+/g, "-") + ".ics";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
 
function categoriesFor(mine) {
  const set = new Set(mine.map((a) => (ROLE_META[a.role] || {}).category).filter(Boolean));
  return CATEGORY_ORDER.filter((c) => set.has(c));
}
function buildDayLabel(row, meta, isToday) {
  const parts = [row.day + " " + row.date];
  if (isToday) parts.push("today");
  if (row.mine.length === 0) parts.push("not listed on rota");
  else parts.push(row.mine.map((a) => a.role).join(", "));
  if (meta && meta.confirmed) parts.push("confirmed");
  if (meta && meta.note) parts.push("has a note");
  return parts.join(", ");
}
 
/* ---------- Hooks ---------- */
function usePrefersReducedMotion() {
  const [reduced, setReduced] = useState(false);
  useEffect(() => {
    const mq = window.matchMedia("(prefers-reduced-motion: reduce)");
    setReduced(mq.matches);
    const handler = (e) => setReduced(e.matches);
    mq.addEventListener("change", handler);
    return () => mq.removeEventListener("change", handler);
  }, []);
  return reduced;
}
function useCountUp(target, duration) {
  const reduced = usePrefersReducedMotion();
  const [value, setValue] = useState(reduced ? target : 0);
  useEffect(() => {
    if (reduced) { setValue(target); return; }
    let raf, start = null;
    const step = (ts) => {
      if (start === null) start = ts;
      const progress = Math.min((ts - start) / (duration || 700), 1);
      setValue(Math.round(progress * target));
      if (progress < 1) raf = requestAnimationFrame(step);
    };
    raf = requestAnimationFrame(step);
    return () => raf && cancelAnimationFrame(raf);
  }, [target, duration, reduced]);
  return value;
}
function useNow() {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  return now;
}
 
/* ---------- Theme ---------- */
const ThemeCtx = createContext({ dark: false });
function useTheme() { return useContext(ThemeCtx); }
 
/* ---------- Next duty lookup ---------- */
function findNextDuty(rows, today) {
  if (today == null) return null;
  const upcoming = rows.filter((r) => r.date >= today && r.mine.length > 0);
  return upcoming.length ? upcoming[0] : null;
}
 
/* ---------- Lightweight toast system ---------- */
function useToasts() {
  const [toasts, setToasts] = useState([]);
  const push = (message) => {
    const id = Date.now() + Math.random();
    setToasts((t) => [...t, { id, message }]);
    window.setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 2800);
  };
  return { toasts, push };
}
 
function ToastStack({ toasts, dark }) {
  if (toasts.length === 0) return null;
  return (
    <div className="fixed top-4 inset-x-0 z-50 flex flex-col items-center gap-2 pointer-events-none px-4" aria-live="polite" aria-atomic="true">
      {toasts.map((t) => (
        <div
          key={t.id}
          className={
            "toast-in pointer-events-auto max-w-xs text-center rounded-full px-4 py-2 text-sm font-semibold shadow-lg border " +
            (dark ? "bg-slate-800 text-slate-100 border-slate-700" : "bg-slate-900 text-white border-slate-900")
          }
        >
          {t.message}
        </div>
      ))}
    </div>
  );
}
 
/* ---------- Focus trap for modal dialogs ---------- */
function useFocusTrap(containerRef, active) {
  useEffect(() => {
    if (!active || !containerRef.current) return;
    const container = containerRef.current;
    const selector = 'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])';
    const handleKey = (e) => {
      if (e.key !== "Tab") return;
      const focusable = Array.from(container.querySelectorAll(selector)).filter((el) => !el.disabled);
      if (focusable.length === 0) return;
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (e.shiftKey && document.activeElement === first) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && document.activeElement === last) {
        e.preventDefault();
        first.focus();
      }
    };
    container.addEventListener("keydown", handleKey);
    return () => container.removeEventListener("keydown", handleKey);
  }, [active, containerRef]);
}
 
/* ---------- Global styles / keyframes ---------- */
function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');
 
      :root { --ease-smooth: cubic-bezier(0.4, 0, 0.2, 1); }
 
      .ecg-path { stroke-dasharray: 5 5; animation: ecgScroll 1s linear infinite; }
      @keyframes ecgScroll { to { stroke-dashoffset: -20; } }
 
      @keyframes cellIn { from { opacity: 0; transform: scale(0.85); } to { opacity: 1; transform: scale(1); } }
      .cell-in { animation: cellIn 0.4s var(--ease-smooth) both; }
 
      @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
      .modal-backdrop { animation: fadeIn 0.25s var(--ease-smooth) both; }
      .panel-fade { animation: fadeIn 0.3s var(--ease-smooth) both; }
 
      @keyframes slideUp { from { transform: translateY(28px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
      .modal-sheet { animation: slideUp 0.35s var(--ease-smooth) both; }
 
      @keyframes slideDown { from { transform: translateY(-10px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
      .panel-drop { animation: slideDown 0.35s var(--ease-smooth) both; }
 
      @keyframes savedPop { 0% { opacity: 0; transform: scale(0.8); } 40% { opacity: 1; transform: scale(1.08); } 100% { opacity: 1; transform: scale(1); } }
      .saved-pop { animation: savedPop 0.3s var(--ease-smooth) both; }
 
      @keyframes blink { 0%, 45% { opacity: 1; } 50%, 95% { opacity: 0.25; } 100% { opacity: 1; } }
      .blink-colon { animation: blink 1.8s step-end infinite; }
 
      @keyframes tabFade { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
      .tab-fade { animation: tabFade 0.3s var(--ease-smooth) both; }
 
      @keyframes toastIn { from { opacity: 0; transform: translateY(-10px) scale(0.95); } to { opacity: 1; transform: translateY(0) scale(1); } }
      .toast-in { animation: toastIn 0.3s var(--ease-smooth) both; }
 
      @keyframes iconFlip { from { opacity: 0; transform: rotate(-90deg) scale(0.5); } to { opacity: 1; transform: rotate(0) scale(1); } }
      .icon-flip { animation: iconFlip 0.35s var(--ease-smooth) both; }
 
      @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
      .shimmer { background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent); background-size: 200% 100%; animation: shimmer 1.4s linear infinite; }
 
      .check-draw path { stroke-dasharray: 20; stroke-dashoffset: 20; animation: checkDraw 0.35s var(--ease-smooth) 0.05s forwards; }
      @keyframes checkDraw { to { stroke-dashoffset: 0; } }
 
      .theme-fade { transition: background-color 0.35s var(--ease-smooth), border-color 0.35s var(--ease-smooth), color 0.35s var(--ease-smooth); }
 
      .header-glow { background-size: 200% 200%; animation: gradientShift 12s ease infinite; }
      @keyframes gradientShift { 0%, 100% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } }
 
      @media print {
        .no-print { display: none !important; }
        .print-only { display: block !important; }
        body { background: white !important; }
      }
 
      @media (prefers-reduced-motion: reduce) {
        .ecg-path, .cell-in, .modal-backdrop, .modal-sheet, .saved-pop, .blink-colon,
        .panel-drop, .panel-fade, .tab-fade, .toast-in, .icon-flip, .shimmer,
        .check-draw path, .header-glow {
          animation: none !important;
        }
        .theme-fade { transition: none !important; }
      }
    `}</style>
  );
}
 
function PulseTrace({ colorClass }) {
  return (
    <svg viewBox="0 0 300 40" preserveAspectRatio="none" className={"w-full h-8 " + (colorClass || "text-teal-300/70")} aria-hidden="true">
      <path
        d="M0 22 L54 22 L64 22 L72 4 L82 36 L90 14 L98 22 L140 22 L150 22 L158 4 L168 36 L176 14 L184 22 L300 22"
        fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="ecg-path"
      />
    </svg>
  );
}
 
function ShiftIcon({ shift, className }) {
  const cls = className || "w-3.5 h-3.5";
  if (shift === "night") return <Moon className={cls} />;
  if (shift === "24h") return <Clock className={cls} />;
  return <Sun className={cls} />;
}
 
function DrawnCheck({ className }) {
  return (
    <svg viewBox="0 0 24 24" className={"check-draw " + (className || "w-4 h-4")} fill="none" aria-hidden="true">
      <path d="M4 12l5 5L20 6" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}
 
function RoleRow({ role }) {
  const { dark } = useTheme();
  const meta = ROLE_META[role] || { time: "", shift: "day", category: "theatre" };
  const style = CATEGORY_STYLE[meta.category];
  const s = dark ? style.dark : style.light;
  return (
    <div className={"theme-fade flex items-center gap-3 rounded-2xl border px-3 py-2.5 " + s.chipBg + " " + s.chipBorder}>
      <div className={"w-8 h-8 shrink-0 rounded-full flex items-center justify-center text-white " + s.solid}>
        <ShiftIcon shift={meta.shift} />
      </div>
      <div className="flex-1 min-w-0">
        <div className={"text-sm font-semibold " + s.chipText}>{role}</div>
        <div className={dark ? "text-xs text-slate-400" : "text-xs text-slate-500"}>{meta.time}</div>
      </div>
    </div>
  );
}
 
function Legend() {
  const { dark } = useTheme();
  return (
    <div className="flex flex-wrap gap-x-3 gap-y-1.5 px-5 mt-3" role="group" aria-label="Color legend">
      {LEGEND_ITEMS.map(([key, label]) => (
        <div key={key} className="flex items-center gap-1.5">
          <span className={"w-2 h-2 rounded-full " + CATEGORY_STYLE[key].dot} aria-hidden="true" />
          <span className={dark ? "text-xs text-slate-400" : "text-xs text-slate-500"}>{label}</span>
        </div>
      ))}
    </div>
  );
}
 
function StatCard({ value, suffix, label, tone, icon }) {
  const { dark } = useTheme();
  const shown = useCountUp(value);
  const toneClasses = dark
    ? { slate: "bg-slate-900 border-slate-700 text-slate-100", amber: "bg-amber-950 border-amber-800 text-amber-300", sky: "bg-sky-950 border-sky-800 text-sky-300" }[tone]
    : { slate: "bg-white border-slate-200 text-slate-800", amber: "bg-amber-50 border-amber-200 text-amber-800", sky: "bg-sky-50 border-sky-200 text-sky-800" }[tone];
  return (
    <div className={"theme-fade rounded-2xl border p-3 text-center transition-transform hover:-translate-y-0.5 " + toneClasses}>
      <div className="flex items-center justify-center gap-1">
        {icon}
        <div className="text-2xl font-bold" style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
          {shown}
          {suffix && <span className="text-sm font-normal opacity-60">{suffix}</span>}
        </div>
      </div>
      <div className="text-xs mt-0.5 leading-tight opacity-80">{label}</div>
    </div>
  );
}
 
function LiveClock({ now }) {
  const dateStr = now.toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const h = pad2(now.getHours());
  const m = pad2(now.getMinutes());
  const s = pad2(now.getSeconds());
  return (
    <div className="flex items-center justify-between mt-3 rounded-2xl bg-white/10 backdrop-blur-sm px-3.5 py-2.5">
      <div className="flex items-center gap-1.5">
        <span className="relative flex h-1.5 w-1.5">
          <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75" />
          <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-400" />
        </span>
        <span className="text-xs sm:text-sm text-teal-50/90">{dateStr}</span>
      </div>
      <span className="font-bold text-white text-base tabular-nums" style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
        {h}<span className="blink-colon">:</span>{m}<span className="blink-colon">:</span>{s}
      </span>
    </div>
  );
}
 
function DarkToggle({ dark, onToggle }) {
  return (
    <button
      type="button" onClick={onToggle} aria-pressed={dark} aria-label={dark ? "Switch to light mode" : "Switch to dark mode"}
      className="w-9 h-9 rounded-full flex items-center justify-center bg-white/10 text-white hover:bg-white/20 active:scale-90 transition-transform focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white overflow-hidden"
    >
      <span key={dark ? "sun" : "moon"} className="icon-flip flex items-center justify-center">
        {dark ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
      </span>
    </button>
  );
}
 
function SkeletonBlock({ className, dark }) {
  return <div className={"rounded-xl relative overflow-hidden " + (dark ? "bg-slate-800" : "bg-slate-200") + " " + className}><span className="shimmer absolute inset-0" /></div>;
}
 
/* ---------- Next duty callout ---------- */
function NextDutyCard({ row, today, monthLabel }) {
  const { dark } = useTheme();
  const base = dark ? "bg-gradient-to-r from-teal-950 to-slate-900 border-teal-800" : "bg-gradient-to-r from-teal-50 to-emerald-50 border-teal-200";
  if (!row) {
    return (
      <div className={"theme-fade rounded-2xl border px-4 py-3 " + base}>
        <p className={"text-sm font-medium " + (dark ? "text-slate-300" : "text-slate-600")}>
          No more duty days left in {monthLabel}.
        </p>
      </div>
    );
  }
  const daysAway = row.date - today;
  const when = daysAway === 0 ? "Today" : daysAway === 1 ? "Tomorrow" : "In " + daysAway + " days";
  const cats = categoriesFor(row.mine);
  return (
    <div className={"theme-fade rounded-2xl border px-4 py-3 " + base}>
      <div className="flex items-center justify-between">
        <div>
          <p className={"text-xs font-semibold uppercase tracking-wide " + (dark ? "text-teal-400" : "text-teal-700")}>Next duty · {when}</p>
          <p className={"text-sm font-bold mt-0.5 " + (dark ? "text-slate-100" : "text-slate-800")} style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
            {row.day} {row.date} — {row.mine.map((a) => a.role).join(", ")}
          </p>
        </div>
        <div className="flex gap-1 shrink-0 ml-2">
          {cats.slice(0, 3).map((c) => <span key={c} className={"w-2 h-2 rounded-full " + CATEGORY_STYLE[c].dot} aria-hidden="true" />)}
        </div>
      </div>
    </div>
  );
}
 
/* ---------- Calendar grid (month view) ---------- */
function CalendarGrid({ rota, rows, dayMeta, today, onSelect, matches }) {
  const { dark } = useTheme();
  if (!rows.length) return null;
  const firstDayIndex = DAY_COLS.indexOf(rows[0].day);
  const leadingBlanks = Array.from({ length: Math.max(firstDayIndex, 0) });
  return (
    <div>
      <div className="grid grid-cols-7 gap-1 mb-1.5" aria-hidden="true">
        {DAY_COLS.map((d) => (
          <div key={d} className={"text-center text-xs font-semibold py-1 " + (d === "FRI" || d === "SAT" ? (dark ? "text-teal-400" : "text-teal-600") : dark ? "text-slate-500" : "text-slate-400")}>
            {d[0]}
          </div>
        ))}
      </div>
      <div className="grid grid-cols-7 gap-1.5" role="grid" aria-label={"Calendar for " + rota.monthLabel}>
        {leadingBlanks.map((_, i) => <div key={"b" + i} aria-hidden="true" />)}
        {rows.map((r, idx) => {
          const cats = categoriesFor(r.mine);
          const meta = dayMeta[dateKeyFor(rota.year, rota.monthIndex, r.date)];
          const isToday = today === r.date;
          const isDimmed = !matches(r);
          const label = buildDayLabel(r, meta, isToday);
          const base = dark ? "border-slate-700 bg-slate-900 hover:border-teal-600 hover:shadow-md" : "border-slate-200 bg-white hover:border-teal-300 hover:shadow-md";
          const todayCls = dark ? "border-teal-500 bg-teal-950" : "border-teal-500 bg-teal-50";
          return (
            <button
              key={r.date} type="button" onClick={() => onSelect(r.date)} aria-label={label}
              className={
                "cell-in theme-fade relative aspect-square rounded-xl border flex flex-col items-center justify-center gap-1 transition-all duration-150 active:scale-90 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 focus-visible:ring-offset-1 " +
                (isToday ? todayCls : base) + (isDimmed ? " opacity-30" : "")
              }
              style={{ animationDelay: idx * 12 + "ms" }}
            >
              {isToday && <span className="absolute inset-0 rounded-xl bg-teal-400 animate-ping opacity-20" aria-hidden="true" />}
              <span className={"text-sm font-semibold relative " + (isToday ? (dark ? "text-teal-300" : "text-teal-700") : dark ? "text-slate-200" : "text-slate-700")}>{r.date}</span>
              <span className="flex gap-0.5 h-1.5 relative">
                {cats.slice(0, 4).map((c) => <span key={c} className={"w-1.5 h-1.5 rounded-full " + CATEGORY_STYLE[c].dot} aria-hidden="true" />)}
              </span>
              {meta && meta.confirmed && <Check className="absolute top-0.5 right-0.5 w-3 h-3 text-emerald-500" strokeWidth={3} aria-hidden="true" />}
              {meta && meta.note && <span className={"absolute bottom-1 left-1 w-1.5 h-1.5 rounded-full " + (dark ? "bg-slate-500" : "bg-slate-400")} aria-hidden="true" />}
            </button>
          );
        })}
      </div>
    </div>
  );
}
 
/* ---------- Agenda list (list view) ---------- */
function AgendaList({ rota, rows, dayMeta, today, onSelect, matches }) {
  const { dark } = useTheme();
  const visible = rows.filter(matches);
  if (visible.length === 0) {
    return <p className={"text-sm italic px-1 " + (dark ? "text-slate-500" : "text-slate-400")}>No days match this filter.</p>;
  }
  return (
    <div className="space-y-2">
      {visible.map((r) => {
        const isToday = today === r.date;
        const meta = dayMeta[dateKeyFor(rota.year, rota.monthIndex, r.date)];
        const cats = categoriesFor(r.mine);
        const stripe = cats[0] ? CATEGORY_STYLE[cats[0]].dot : dark ? "bg-slate-700" : "bg-slate-200";
        return (
          <button
            key={r.date} type="button" onClick={() => onSelect(r.date)} aria-label={buildDayLabel(r, meta, isToday)}
            className={
              "theme-fade w-full text-left flex rounded-xl border overflow-hidden transition-all active:scale-[0.99] hover:shadow-md focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " +
              (dark ? "bg-slate-900 " : "bg-white ") + (isToday ? "border-teal-400 ring-1 ring-teal-800" : dark ? "border-slate-700" : "border-slate-200")
            }
          >
            <div className={"w-1.5 " + stripe} />
            <div className="flex-1 p-3">
              <div className="flex items-center justify-between mb-1.5">
                <div className="flex items-center gap-2">
                  <span className={"font-bold " + (dark ? "text-slate-100" : "text-slate-800")} style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
                    {r.day} {r.date}
                  </span>
                  {meta && meta.note && <StickyNote className={"w-3 h-3 " + (dark ? "text-slate-500" : "text-slate-400")} aria-hidden="true" />}
                  {meta && meta.confirmed && <Check className="w-3 h-3 text-emerald-500" strokeWidth={3} aria-hidden="true" />}
                </div>
                {isToday && <span className={"text-xs font-semibold rounded-full px-2 py-0.5 " + (dark ? "text-teal-300 bg-teal-950" : "text-teal-700 bg-teal-50")}>TODAY</span>}
              </div>
              {r.mine.length === 0 ? (
                <div className={"text-sm italic " + (dark ? "text-slate-500" : "text-slate-400")}>Not listed on rota</div>
              ) : (
                <div className="flex flex-wrap gap-1">
                  {r.mine.map((a, i) => {
                    const style = CATEGORY_STYLE[(ROLE_META[a.role] || {}).category] || CATEGORY_STYLE.theatre;
                    const s = dark ? style.dark : style.light;
                    return <span key={i} className={"text-xs font-medium rounded-full px-2 py-0.5 border " + s.chipBg + " " + s.chipBorder + " " + s.chipText}>{a.role}</span>;
                  })}
                </div>
              )}
            </div>
          </button>
        );
      })}
    </div>
  );
}
 
/* ---------- Day detail modal ---------- */
function DayModal({ row, meta, today, onClose, onSaveNote, onToggleConfirm }) {
  const { dark } = useTheme();
  const [draft, setDraft] = useState((meta && meta.note) || "");
  const [savedFlash, setSavedFlash] = useState(false);
  const closeBtnRef = useRef(null);
  const panelRef = useRef(null);
  useFocusTrap(panelRef, true);
 
  useEffect(() => { setDraft((meta && meta.note) || ""); }, [row.date]);
 
  useEffect(() => {
    closeBtnRef.current && closeBtnRef.current.focus();
    const onKey = (e) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);
 
  const handleSave = () => {
    onSaveNote(row.date, draft);
    setSavedFlash(true);
    window.setTimeout(() => setSavedFlash(false), 1600);
  };
 
  const isToday = today === row.date;
  const confirmed = !!(meta && meta.confirmed);
  const panelBg = dark ? "bg-slate-900" : "bg-white";
  const borderCol = dark ? "border-slate-700" : "border-slate-100";
  const textPrimary = dark ? "text-slate-100" : "text-slate-800";
  const textMuted = "text-slate-400";
 
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center">
      <div className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm modal-backdrop" onClick={onClose} />
      <div
        ref={panelRef} role="dialog" aria-modal="true" aria-labelledby="day-modal-title"
        className={"relative w-full sm:max-w-sm sm:rounded-3xl rounded-t-3xl overflow-y-auto modal-sheet theme-fade " + panelBg}
        style={{ maxHeight: "85vh" }}
      >
        <div className={"sticky top-0 flex items-center justify-between px-5 pt-5 pb-3 border-b theme-fade " + panelBg + " " + borderCol}>
          <div>
            <div id="day-modal-title" className={"text-lg font-bold " + textPrimary} style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
              {row.day} {row.date}
            </div>
            {isToday && <span className={dark ? "text-xs font-semibold text-teal-400" : "text-xs font-semibold text-teal-700"}>Today</span>}
          </div>
          <button
            ref={closeBtnRef} type="button" onClick={onClose} aria-label="Close day details"
            className={"w-9 h-9 rounded-full flex items-center justify-center transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " + (dark ? "text-slate-400 hover:bg-slate-800 hover:text-slate-200" : "text-slate-400 hover:bg-slate-100 hover:text-slate-600")}
          >
            <X className="w-5 h-5" />
          </button>
        </div>
 
        <div className="px-5 py-4 space-y-4">
          <div>
            <h3 className={"text-xs font-semibold uppercase tracking-widest mb-2 " + textMuted}>Assignments</h3>
            {row.mine.length === 0 ? (
              <p className={"text-sm italic " + textMuted}>Not listed on rota for this day.</p>
            ) : (
              <div className="space-y-2">{row.mine.map((a, i) => <RoleRow key={i} role={a.role} />)}</div>
            )}
          </div>
 
          <button
            type="button" onClick={() => onToggleConfirm(row.date)} aria-pressed={confirmed}
            className={
              "w-full flex items-center justify-center gap-2 rounded-xl border py-2.5 text-sm font-semibold transition-all active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " +
              (confirmed
                ? dark ? "bg-emerald-950 border-emerald-700 text-emerald-300" : "bg-emerald-50 border-emerald-300 text-emerald-700"
                : dark ? "bg-slate-800 border-slate-700 text-slate-400 hover:border-slate-600" : "bg-white border-slate-200 text-slate-500 hover:border-slate-300")
            }
          >
            {confirmed ? <DrawnCheck className="w-4 h-4" /> : <Check className="w-4 h-4" strokeWidth={3} />}
            {confirmed ? "Confirmed" : "Mark as confirmed"}
          </button>
 
          <div>
            <label htmlFor="day-note" className={"text-xs font-semibold uppercase tracking-widest mb-2 block " + textMuted}>Your note</label>
            <textarea
              id="day-note" value={draft} onChange={(e) => setDraft(e.target.value)} rows={3}
              placeholder="Swap requested, bring extra scrubs, etc."
              className={"w-full rounded-xl border p-3 text-sm resize-none transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " + (dark ? "bg-slate-800 border-slate-700 text-slate-100 placeholder:text-slate-500" : "bg-white border-slate-200 text-slate-700 placeholder:text-slate-400")}
            />
            <div className="flex items-center gap-3 mt-2">
              <button type="button" onClick={handleSave} className="rounded-xl bg-teal-700 text-white text-sm font-semibold px-4 py-2 hover:bg-teal-800 active:scale-95 transition-transform focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 focus-visible:ring-offset-1">
                Save note
              </button>
              {savedFlash && (
                <span className="saved-pop text-xs font-semibold text-emerald-500 flex items-center gap-1">
                  <Check className="w-3.5 h-3.5" strokeWidth={3} /> Saved
                </span>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
 
/* ---------- Upload panel: bring in a new month's rota ---------- */
function UploadPanel({ onLoaded, onClose }) {
  const { dark } = useTheme();
  const [status, setStatus] = useState("idle");
  const [error, setError] = useState("");
  const [parsed, setParsed] = useState(null);
  const [source, setSource] = useState("xlsx");
  const [monthGuess, setMonthGuess] = useState("");
  const fileRef = useRef(null);
 
  const handleFile = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setStatus("parsing");
    setError("");
    try {
      const isPdf = /\.pdf$/i.test(file.name);
      const isXlsx = /\.xlsx?$/i.test(file.name);
      if (!isPdf && !isXlsx) {
        throw new Error("Please choose the Master Duty Rota as a .xlsx or .pdf file.");
      }
 
      if (isPdf) {
        setSource("pdf");
        setStatus("parsing-pdf");
        const grid = await extractGridFromPdf(file);
        const result = parseGrid(grid);
        if (result.days.length === 0) {
          throw new Error("Read the PDF but couldn't line up any day rows. Try the .xlsx file instead, or send the PDF to Claude in chat.");
        }
        setParsed(result);
        setMonthGuess(detectMonthYear(grid));
        setStatus("preview");
        return;
      }
 
      setSource("xlsx");
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      if (!wb.SheetNames || wb.SheetNames.length === 0) {
        throw new Error("This workbook doesn't seem to have any sheets.");
      }
      let result = null;
      let usedGrid = null;
      for (const sheetName of wb.SheetNames) {
        const sheet = wb.Sheets[sheetName];
        const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const r = parseGrid(grid);
        if (r.days.length > 0) { result = r; usedGrid = grid; break; }
      }
      if (!result) {
        throw new Error("Couldn't find day rows in any sheet of this file. Make sure it's the Master Duty Rota, same layout as usual.");
      }
      setParsed(result);
      setMonthGuess(detectMonthYear(usedGrid));
      setStatus("preview");
    } catch (err) {
      setError(err && err.message ? err.message : "Couldn't read that file.");
      setStatus("error");
    }
  };
 
  const confirmLoad = () => {
    if (!monthGuess) { setError("Pick which month this rota is for."); return; }
    const [y, m] = monthGuess.split("-").map(Number);
    onLoaded({ monthLabel: MONTH_NAMES[m - 1] + " " + y, year: y, monthIndex: m - 1, days: parsed.days, names: parsed.names });
  };
 
  const inputCls = dark ? "bg-slate-800 border-slate-700 text-slate-100" : "bg-white border-slate-200 text-slate-800";
  const sampleDays = parsed ? parsed.days.filter((d) => d.assignments.length > 0).slice(0, source === "pdf" ? 3 : 1) : [];
 
  return (
    <div className={"panel-drop theme-fade rounded-2xl border p-4 mt-3 " + (dark ? "bg-slate-900 border-slate-700" : "bg-white border-slate-200")}>
      <div className="flex items-center justify-between mb-2">
        <h3 className={"text-sm font-bold " + (dark ? "text-slate-100" : "text-slate-800")}>Load a new month</h3>
        <button type="button" onClick={onClose} aria-label="Close upload panel" className={dark ? "text-slate-500 hover:text-slate-300" : "text-slate-400 hover:text-slate-600"}>
          <X className="w-4 h-4" />
        </button>
      </div>
 
      {status !== "preview" && (
        <>
          <p className={"text-xs mb-3 " + (dark ? "text-slate-400" : "text-slate-500")}>
            Upload the Master Duty Rota for the next month. <strong>.xlsx</strong> is read directly and is the most reliable. <strong>.pdf</strong> is also accepted — it's read with a best-effort column matcher, so double-check the preview before loading.
          </p>
          <label className={"flex items-center justify-center gap-2 rounded-xl border-2 border-dashed py-4 cursor-pointer text-sm font-semibold transition-colors " + (dark ? "border-slate-700 text-slate-300 hover:border-teal-600" : "border-slate-300 text-slate-600 hover:border-teal-400")}>
            {(status === "parsing" || status === "parsing-pdf") ? <Loader2 className="w-4 h-4 animate-spin" /> : <Upload className="w-4 h-4" />}
            {status === "parsing-pdf" ? "Reading PDF… this can take a moment" : status === "parsing" ? "Reading file…" : "Choose .xlsx or .pdf file"}
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.pdf" className="hidden" onChange={handleFile} disabled={status === "parsing" || status === "parsing-pdf"} />
          </label>
        </>
      )}
 
      {status === "error" && (
        <div role="alert" className={"flex items-start gap-2 mt-3 text-xs rounded-lg p-2.5 " + (dark ? "bg-red-950 text-red-300" : "bg-red-50 text-red-700")}>
          <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
          <span>{error}</span>
        </div>
      )}
 
      {status === "preview" && parsed && (
        <div className="tab-fade">
          <p className={"text-sm mb-2 " + (dark ? "text-slate-300" : "text-slate-600")}>
            Found <strong>{parsed.days.length}</strong> days and <strong>{parsed.names.length}</strong> staff names.
          </p>
 
          {source === "pdf" && (
            <div className={"flex items-start gap-2 rounded-xl border p-2.5 mb-3 text-xs " + (dark ? "bg-amber-950 border-amber-800 text-amber-300" : "bg-amber-50 border-amber-300 text-amber-800")}>
              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
              <span>Read from a PDF, so columns were matched by position rather than read directly — less certain than Excel. Check all three days below carefully against the actual file before loading.</span>
            </div>
          )}
 
          {sampleDays.map((d) => (
            <div key={d.date} className={"rounded-xl border p-2.5 mb-2 text-xs " + (dark ? "bg-slate-800 border-slate-700 text-slate-300" : "bg-slate-50 border-slate-200 text-slate-600")}>
              <span className={"font-semibold " + (dark ? "text-slate-200" : "text-slate-700")}>Quick check — {d.day} {d.date}: </span>
              {d.assignments.length
                ? d.assignments.slice(0, 4).map((a) => a.role + " (" + a.name + ")").join(", ") + (d.assignments.length > 4 ? "…" : "")
                : "no assignments found"}
            </div>
          ))}
          <p className={"text-xs mb-3 opacity-80 " + (dark ? "text-slate-400" : "text-slate-500")}>
            Does this look right? If not, cancel and try the other file format, or send it to Claude in chat.
          </p>
 
          <label className="block mb-3">
            <span className={"text-xs mb-1 block " + (dark ? "text-slate-400" : "text-slate-500")}>Confirm which month this is</span>
            <input
              type="month" value={monthGuess} onChange={(e) => setMonthGuess(e.target.value)}
              className={"w-full rounded-xl px-3 py-2 border text-base focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " + inputCls}
            />
          </label>
          {error && (
            <div role="alert" className={"flex items-start gap-2 mb-3 text-xs rounded-lg p-2.5 " + (dark ? "bg-red-950 text-red-300" : "bg-red-50 text-red-700")}>
              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" /><span>{error}</span>
            </div>
          )}
          <div className="flex gap-2">
            <button type="button" onClick={confirmLoad} className="flex-1 rounded-xl bg-teal-700 text-white text-sm font-semibold py-2.5 hover:bg-teal-800 active:scale-[0.98] transition-transform focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500">
              Load this rota
            </button>
            <button
              type="button"
              onClick={() => { setStatus("idle"); setParsed(null); setError(""); if (fileRef.current) fileRef.current.value = ""; }}
              className={"rounded-xl border text-sm font-semibold px-4 py-2.5 transition-colors " + (dark ? "border-slate-700 text-slate-300" : "border-slate-200 text-slate-600")}
            >
              Start over
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
 
/* ---------- Loading skeleton (shown briefly while local storage hydrates) ---------- */
function SkeletonScreen({ dark }) {
  return (
    <div className={"min-h-screen " + (dark ? "bg-slate-950" : "bg-slate-50")}>
      <div className="max-w-md mx-auto pb-10 px-5 pt-6">
        <SkeletonBlock dark={dark} className="h-40 w-full mb-4" />
        <div className="grid grid-cols-3 gap-2 mb-3">
          <SkeletonBlock dark={dark} className="h-16" />
          <SkeletonBlock dark={dark} className="h-16" />
          <SkeletonBlock dark={dark} className="h-16" />
        </div>
        <SkeletonBlock dark={dark} className="h-64 w-full" />
      </div>
    </div>
  );
}
 
/* ---------- Main app ---------- */
export default function AnesthesiaMonthlyRota() {
  const [rota, setRota] = useState(DEFAULT_ROTA);
  const [contacts] = useState(DEFAULT_CONTACTS);
  const [name, setName] = useState("");
  const [dark, setDark] = useState(false);
  const [dayMeta, setDayMeta] = useState({});
  const [tab, setTab] = useState("month");
  const [filter, setFilter] = useState("all");
  const [selectedDate, setSelectedDate] = useState(null);
  const [showUpload, setShowUpload] = useState(false);
  const [hydrated, setHydrated] = useState(false);
  const now = useNow();
  const { toasts, push } = useToasts();
 
  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("rota-current-data", false);
        if (r && r.value) setRota(JSON.parse(r.value));
      } catch (e) {}
      try {
        const n = await window.storage.get("rota-selected-name", false);
        if (n && n.value) setName(n.value);
      } catch (e) {}
      try {
        const mt = await window.storage.get("rota-day-meta", false);
        if (mt && mt.value) setDayMeta(JSON.parse(mt.value));
      } catch (e) {}
      try {
        const dm = await window.storage.get("rota-dark-mode", false);
        if (dm && dm.value === "1") setDark(true);
      } catch (e) {}
      setHydrated(true);
    })();
  }, []);
 
  useEffect(() => { if (hydrated) window.storage.set("rota-selected-name", name, false).catch(() => {}); }, [name, hydrated]);
  useEffect(() => { if (hydrated) window.storage.set("rota-dark-mode", dark ? "1" : "0", false).catch(() => {}); }, [dark, hydrated]);
 
  const persistMeta = async (next) => {
    setDayMeta(next);
    try { await window.storage.set("rota-day-meta", JSON.stringify(next), false); } catch (e) {}
  };
  const handleSaveNote = (date, note) => {
    const key = dateKeyFor(rota.year, rota.monthIndex, date);
    persistMeta({ ...dayMeta, [key]: { ...dayMeta[key], note } });
  };
  const handleToggleConfirm = (date) => {
    const key = dateKeyFor(rota.year, rota.monthIndex, date);
    const cur = !!(dayMeta[key] && dayMeta[key].confirmed);
    persistMeta({ ...dayMeta, [key]: { ...dayMeta[key], confirmed: !cur } });
  };
  const handleRotaLoaded = async (newRota) => {
    setRota(newRota);
    setShowUpload(false);
    setTab("month");
    setFilter("all");
    setSelectedDate(null);
    if (newRota.names.indexOf(name) === -1) setName("");
    try { await window.storage.set("rota-current-data", JSON.stringify(newRota), false); } catch (e) {}
    push("Loaded " + newRota.monthLabel + " rota");
  };
  const handlePrint = () => {
    push("Opening print dialog…");
    window.setTimeout(() => window.print(), 150);
  };
  const handleExportICS = () => {
    downloadICS(rota, name, dayMeta);
    push("Calendar file downloaded");
  };
 
  const rows = useMemo(() => rota.days.map((d) => ({ ...d, mine: d.assignments.filter((a) => a.name === name) })), [rota, name]);
 
  const stats = useMemo(() => {
    if (!name) return null;
    const dutyDays = rows.filter((r) => r.mine.length > 0).length;
    const businessDays = rows.filter((r) => r.mine.some((a) => a.role === "Business Coverage")).length;
    const saturdays = rows.filter((r) => r.day === "SAT");
    const saturdayDuty = saturdays.filter((r) => r.mine.some((a) => a.role === "Saturday List")).length;
    return { dutyDays, businessDays, saturdayDuty, totalSaturdays: saturdays.length };
  }, [rows, name]);
 
  const today = useMemo(() => {
    if (now.getFullYear() === rota.year && now.getMonth() === rota.monthIndex) return now.getDate();
    return null;
  }, [now, rota.year, rota.monthIndex]);
 
  const nextDutyRow = useMemo(() => (name ? findNextDuty(rows, today) : null), [rows, today, name]);
 
  const matches = (row) => {
    if (filter === "all") return true;
    if (filter === "business") return row.mine.some((a) => a.role === "Business Coverage");
    if (filter === "saturday") return row.mine.some((a) => a.role === "Saturday List");
    if (filter === "noted") {
      const key = dateKeyFor(rota.year, rota.monthIndex, row.date);
      return !!(dayMeta[key] && dayMeta[key].note);
    }
    return true;
  };
 
  const selectedRow = selectedDate != null ? rows.find((r) => r.date === selectedDate) : null;
  const selectedMeta = selectedRow ? dayMeta[dateKeyFor(rota.year, rota.monthIndex, selectedRow.date)] : null;
 
  if (!hydrated) return <SkeletonScreen dark={dark} />;
 
  const pageBg = dark ? "bg-slate-950" : "bg-slate-50";
  const exportBtnCls =
    "flex-1 flex items-center justify-center gap-1.5 rounded-xl border text-xs font-semibold py-2.5 active:scale-95 transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " +
    (dark ? "border-slate-700 bg-slate-900 text-slate-200 hover:border-teal-600" : "border-slate-200 bg-white text-slate-700 hover:border-teal-300");
 
  return (
    <ThemeCtx.Provider value={{ dark }}>
      <div className={"min-h-screen theme-fade " + pageBg} style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
        <GlobalStyle />
 
        <div className="no-print max-w-md mx-auto pb-10">
          <div className="header-glow bg-gradient-to-br from-emerald-900 via-teal-800 to-teal-700 text-white px-5 pt-6 pb-5 rounded-b-3xl relative overflow-hidden panel-fade">
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="text-xs font-semibold tracking-widest uppercase text-teal-200" style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
                  Anesthesia Technology
                </div>
                <h1 className="text-xl font-bold mt-0.5 leading-tight" style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
                  Monthly Rota
                </h1>
              </div>
              <DarkToggle dark={dark} onToggle={() => setDark((d) => !d)} />
            </div>
 
            <LiveClock now={now} />
 
            <div className="mt-3 -mx-1">
              <PulseTrace />
            </div>
 
            <div className="flex items-center justify-between gap-2 mt-1 flex-wrap">
              <span className="inline-flex items-center rounded-full bg-white text-teal-800 text-sm font-bold px-3 py-1" style={{ fontFamily: "'Space Grotesk', system-ui, sans-serif" }}>
                {rota.monthLabel} Rota
              </span>
              <button
                type="button" onClick={() => setShowUpload((v) => !v)} aria-pressed={showUpload}
                className="flex items-center gap-1.5 text-xs font-semibold rounded-full bg-white/10 hover:bg-white/20 text-white px-3 py-1.5 active:scale-95 transition-transform focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white"
              >
                <Upload className="w-3.5 h-3.5" /> New month
              </button>
            </div>
 
            <label className="block mt-3">
              <span className="text-xs text-teal-200 mb-1 block">Your name</span>
              <select
                value={name} onChange={(e) => setName(e.target.value)}
                className="w-full rounded-xl px-3 py-2.5 text-slate-900 bg-white text-base focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-300"
              >
                <option value="" disabled>Select your name…</option>
                {rota.names.map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
            </label>
          </div>
 
          {showUpload && (
            <div className="px-5">
              <UploadPanel onLoaded={handleRotaLoaded} onClose={() => setShowUpload(false)} />
            </div>
          )}
 
          {!name && (
            <div className="px-5 pt-10 text-center panel-fade">
              <div className="relative w-14 h-14 mx-auto mb-4">
                <span className="absolute inset-0 rounded-full bg-teal-400/20 animate-ping" />
                <div className={"relative w-14 h-14 rounded-full flex items-center justify-center " + (dark ? "bg-slate-900 border border-slate-700" : "bg-white border border-slate-200")}>
                  <CalendarDays className={"w-6 h-6 " + (dark ? "text-teal-400" : "text-teal-600")} />
                </div>
              </div>
              <p className={"text-sm " + (dark ? "text-slate-400" : "text-slate-500")}>
                Pick your name above to see your {rota.monthLabel} schedule, business coverage days, and Saturday list days.
              </p>
            </div>
          )}
 
          {name && stats && (
            <>
              <div className="px-5 mt-5">
                <NextDutyCard row={nextDutyRow} today={today} monthLabel={rota.monthLabel} />
              </div>
 
              <div className="grid grid-cols-3 gap-2 px-5 mt-3">
                <StatCard value={stats.dutyDays} label="duty days" tone="slate" />
                <StatCard value={stats.businessDays} label="business coverage" tone="amber" />
                <StatCard value={stats.saturdayDuty} suffix={"/" + stats.totalSaturdays} label="Saturday lists" tone="sky" />
              </div>
 
              <Legend />
 
              <div className="px-5 mt-4 flex gap-2">
                <button type="button" onClick={handlePrint} className={exportBtnCls}>
                  <Printer className="w-3.5 h-3.5" /> Export PDF
                </button>
                <button type="button" onClick={handleExportICS} className={exportBtnCls}>
                  <CalendarPlus className="w-3.5 h-3.5" /> Add to Calendar
                </button>
              </div>
 
              <div className="px-5 mt-4">
                <div role="tablist" aria-label="View" className={"flex gap-1.5 rounded-xl p-1 theme-fade " + (dark ? "bg-slate-900" : "bg-slate-100")}>
                  <button
                    type="button" role="tab" aria-selected={tab === "month"} onClick={() => setTab("month")}
                    className={"flex-1 flex items-center justify-center gap-1.5 rounded-lg py-2 text-sm font-semibold transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " +
                      (tab === "month" ? (dark ? "bg-slate-800 text-teal-300 shadow-sm" : "bg-white text-teal-800 shadow-sm") : "text-slate-500")}
                  >
                    <CalendarDays className="w-4 h-4" /> Month
                  </button>
                  <button
                    type="button" role="tab" aria-selected={tab === "list"} onClick={() => setTab("list")}
                    className={"flex-1 flex items-center justify-center gap-1.5 rounded-lg py-2 text-sm font-semibold transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " +
                      (tab === "list" ? (dark ? "bg-slate-800 text-teal-300 shadow-sm" : "bg-white text-teal-800 shadow-sm") : "text-slate-500")}
                  >
                    <ListIcon className="w-4 h-4" /> List
                  </button>
                </div>
              </div>
 
              <div className="px-5 mt-3 flex flex-wrap gap-1.5" role="group" aria-label="Filter days">
                {FILTERS.map((f) => (
                  <button
                    key={f.key} type="button" aria-pressed={filter === f.key} onClick={() => setFilter(f.key)}
                    className={"text-xs font-semibold rounded-full px-3 py-1.5 border transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-500 " +
                      (filter === f.key ? "bg-teal-700 border-teal-700 text-white" : dark ? "bg-slate-900 border-slate-700 text-slate-400" : "bg-white border-slate-200 text-slate-500")}
                  >
                    {f.label}
                  </button>
                ))}
              </div>
 
              <div className="px-5 mt-4 tab-fade" key={tab} role="tabpanel">
                {tab === "month" ? (
                  <CalendarGrid rota={rota} rows={rows} dayMeta={dayMeta} today={today} onSelect={setSelectedDate} matches={matches} />
                ) : (
                  <AgendaList rota={rota} rows={rows} dayMeta={dayMeta} today={today} onSelect={setSelectedDate} matches={matches} />
                )}
              </div>
            </>
          )}
 
          <div className="px-5 mt-8">
            <h2 className={"text-xs font-semibold uppercase tracking-widest mb-2 " + (dark ? "text-slate-500" : "text-slate-400")}>Quick contacts</h2>
            <div className={"theme-fade rounded-xl border divide-y " + (dark ? "bg-slate-900 border-slate-700 divide-slate-800" : "bg-white border-slate-200 divide-slate-100")}>
              {contacts.map((c, i) => (
                <div key={i} className="flex items-center justify-between px-3 py-2.5">
                  <span className={"text-sm " + (dark ? "text-slate-300" : "text-slate-600")}>{c.label}</span>
                  <span className={"flex items-center gap-1 text-sm font-semibold " + (dark ? "text-teal-400" : "text-teal-700")}>
                    <Phone className="w-3.5 h-3.5" />{c.bleep}
                  </span>
                </div>
              ))}
            </div>
          </div>
 
          <p className={"text-center text-xs mt-6 px-5 " + (dark ? "text-slate-600" : "text-slate-400")}>
            {rota.monthLabel} rota · Notes, name, and rota data saved privately on this device.
          </p>
        </div>
 
        {selectedRow && (
          <DayModal row={selectedRow} meta={selectedMeta} today={today} onClose={() => setSelectedDate(null)} onSaveNote={handleSaveNote} onToggleConfirm={handleToggleConfirm} />
        )}
 
        {name && (
          <div className="hidden print-only p-8 text-black bg-white">
            <h1 className="text-2xl font-bold">Anesthesia Technology Monthly Rota</h1>
            <h2 className="text-lg mb-4">{rota.monthLabel} — {name}</h2>
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr>
                  <th className="border border-slate-400 p-2 text-left">Date</th>
                  <th className="border border-slate-400 p-2 text-left">Day</th>
                  <th className="border border-slate-400 p-2 text-left">Assignment</th>
                  <th className="border border-slate-400 p-2 text-left">Note</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r) => {
                  const key = dateKeyFor(rota.year, rota.monthIndex, r.date);
                  const m = dayMeta[key];
                  return (
                    <tr key={r.date}>
                      <td className="border border-slate-400 p-2">{r.date}</td>
                      <td className="border border-slate-400 p-2">{r.day}</td>
                      <td className="border border-slate-400 p-2">{r.mine.length ? r.mine.map((a) => a.role).join(", ") : "—"}</td>
                      <td className="border border-slate-400 p-2">{(m && m.note) || ""}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
 
        <ToastStack toasts={toasts} dark={dark} />
      </div>
    </ThemeCtx.Provider>
  );
}
