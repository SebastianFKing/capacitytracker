import React, { useState, useEffect, useRef } from 'react';
import { 
  Users, 
  LayoutDashboard, 
  Save, 
  Plus, 
  Trash2, 
  LogOut, 
  AlertCircle,
  FileText,
  Lock,
  ArrowLeft,
  ChevronRight,
  ChevronDown,
  Check,
  ArrowUp,
  ArrowDown,
  CalendarDays,
  Settings,
  X
} from 'lucide-react';

// --- Types ---

type Project = {
  id: string;
  name: string;
  category: 'Category A' | 'Category B' | 'Category C';
  owner: string;
  capacities: [number, number, number, number]; // Array of 4 weeks (Percentage 0-100)
  deadline: string;
};

type WeeklyEntry = {
  weekDate: string;
  employeeName: string;
  office: string;
  mentor: string;
  languages: string[]; 
  // 2D array: 4 weeks x 5 days (Mon-Fri)
  annualLeave: boolean[][]; 
  selfAssessment: 'Open' | 'Limited Capacity' | 'At Capacity' | 'Over Capacity';
  projects: Project[];
  lastUpdated: string;
};

type Employee = {
  name: string;
  password: string;
};

type LoginAction =
  | { type: 'employee'; employeeName: string }
  | { type: 'management' }
  | { type: 'operations' }
  | { type: 'it' };

// --- Initial Constants (Default Data) ---

const INITIAL_OFFICES = ['Office A', 'Office B', 'Office C', 'Office D', 'Office E', 'Office F'];
const INITIAL_MENTORS = ['Mentor 1', 'Mentor 2', 'Mentor 3', 'Mentor 4'];
const INITIAL_LANGUAGES = ['English', 'French', 'German', 'Dutch', 'Spanish', 'Mandarin', 'Arabic'];
const PROJECT_CATEGORIES = ['Category A', 'Category B', 'Category C'];
const WEEKDAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];

const INITIAL_EMPLOYEES: Employee[] = [
  { name: 'Employee A', password: 'pass123' },
  { name: 'Employee B', password: 'pass123' },
  { name: 'Employee C', password: 'pass123' },
  { name: 'Employee D', password: 'pass123' },
  { name: 'Employee E', password: 'pass123' },
  { name: 'Employee F', password: 'pass123' }
];

// --- Mock Initial Data ---

const MOCK_DB: Record<string, WeeklyEntry> = {
  'Employee A-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee A',
    office: 'Office A',
    mentor: 'Mentor 2',
    languages: ['English', 'Spanish'],
    // 4 weeks, 5 days each, all false initially
    annualLeave: Array(4).fill(null).map(() => Array(5).fill(false)),
    selfAssessment: 'Limited Capacity',
    lastUpdated: new Date().toISOString(),
    projects: [
      { id: '1', name: 'Task1', category: 'Category A', owner: 'Owner 1', capacities: [25, 25, 20, 10], deadline: 'Deadline1' },
      { id: '2', name: 'Task2', category: 'Category A', owner: 'Owner 2', capacities: [20, 20, 15, 10], deadline: 'Deadline2' },
      { id: '3', name: 'Task3', category: 'Category B', owner: 'Employee A', capacities: [20, 15, 10, 5], deadline: 'Deadline3' },
      { id: '4', name: 'Task4', category: 'Category B', owner: 'Owner 3', capacities: [2, 2, 0, 0], deadline: 'Deadline4' },
      { id: '5', name: 'Task5', category: 'Category B', owner: 'Owner 4', capacities: [15, 10, 5, 0], deadline: 'Deadline5' },
      { id: '6', name: 'Task6', category: 'Category C', owner: 'Owner 5', capacities: [5, 5, 5, 5], deadline: 'Deadline6' },
    ]
  },
  'Employee B-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee B',
    office: 'Office E', 
    mentor: 'Mentor 1',
    languages: ['French'],
    // Example: Mon off in week 1, All week off in week 4
    annualLeave: [
      [true, false, false, false, false],
      [false, false, false, false, false],
      [false, false, false, false, false],
      [true, true, true, true, true]
    ], 
    selfAssessment: 'Over Capacity',
    lastUpdated: new Date().toISOString(),
    projects: [
      { id: '1', name: 'Task7', category: 'Category A', owner: 'Owner 1', capacities: [60, 50, 40, 20], deadline: 'Deadline7' },
      { id: '2', name: 'Task8', category: 'Category B', owner: 'Owner 2', capacities: [50, 50, 50, 10], deadline: 'Deadline8' },
    ]
  },
  'Employee C-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee C',
    office: 'Office F',
    mentor: 'Mentor 1',
    languages: ['English', 'German'],
    annualLeave: [
      [false, false, false, false, false],
      [false, true, false, false, false],
      [false, false, false, false, false],
      [false, false, false, false, false]
    ],
    selfAssessment: 'At Capacity',
    lastUpdated: new Date().toISOString(),
    projects: [
      { id: '1', name: 'Task9', category: 'Category A', owner: 'Owner 6', capacities: [40, 35, 30, 20], deadline: 'Deadline9' },
      { id: '2', name: 'Task10', category: 'Category B', owner: 'Owner 7', capacities: [35, 30, 20, 15], deadline: 'Deadline10' },
      { id: '3', name: 'Task11', category: 'Category C', owner: 'Employee C', capacities: [10, 10, 10, 5], deadline: 'Deadline11' },
    ]
  },
  'Employee D-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee D',
    office: 'Office B',
    mentor: 'Mentor 1',
    languages: ['Spanish', 'English'],
    annualLeave: [
      [false, false, false, false, false],
      [false, false, false, false, false],
      [true, false, false, false, false],
      [false, false, false, false, false]
    ],
    selfAssessment: 'Limited Capacity',
    lastUpdated: new Date().toISOString(),
    projects: [
      { id: '1', name: 'Task12', category: 'Category A', owner: 'Employee D', capacities: [25, 25, 20, 20], deadline: 'Deadline12' },
      { id: '2', name: 'Task13', category: 'Category B', owner: 'Owner 8', capacities: [20, 20, 20, 15], deadline: 'Deadline13' },
      { id: '3', name: 'Task14', category: 'Category C', owner: 'Mentor 2', capacities: [15, 10, 10, 10], deadline: 'Deadline14' },
    ]
  },
  'Employee E-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee E',
    office: 'Office A',
    mentor: 'Mentor 4',
    languages: ['Arabic', 'English', 'French'],
    annualLeave: [
      [false, false, false, false, false],
      [false, false, false, false, false],
      [false, false, false, false, false],
      [false, true, true, false, false]
    ],
    selfAssessment: 'Open',
    lastUpdated: new Date().toISOString(),
    projects: [
      { id: '1', name: 'Task15', category: 'Category A', owner: 'Employee E', capacities: [20, 20, 20, 10], deadline: 'Deadline15' },
      { id: '2', name: 'Task16', category: 'Category B', owner: 'Owner 9', capacities: [15, 20, 20, 15], deadline: 'Deadline16' },
      { id: '3', name: 'Task17', category: 'Category C', owner: 'Mentor 3', capacities: [10, 10, 5, 5], deadline: 'Deadline17' },
    ]
  },
  'Employee F-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee F',
    office: 'Office E',
    mentor: 'Mentor 2',
    languages: ['Mandarin', 'English'],
    annualLeave: [
      [false, false, false, false, false],
      [false, false, false, false, false],
      [false, false, false, false, false],
      [false, false, false, false, false]
    ],
    selfAssessment: 'At Capacity',
    lastUpdated: new Date().toISOString(),
    projects: [
      { id: '1', name: 'Task18', category: 'Category A', owner: 'Employee F', capacities: [30, 35, 35, 25], deadline: 'Deadline18' },
      { id: '2', name: 'Task19', category: 'Category B', owner: 'Owner 2', capacities: [30, 25, 20, 15], deadline: 'Deadline19' },
      { id: '3', name: 'Task20', category: 'Category C', owner: 'Mentor 4', capacities: [15, 15, 10, 10], deadline: 'Deadline20' },
    ]
  }
};

// --- Helper Functions ---

const getWeekLabels = (startDateStr: string) => {
  const startDate = new Date(startDateStr);
  return Array.from({ length: 4 }).map((_, i) => {
    const monday = new Date(startDate);
    monday.setDate(startDate.getDate() + (i * 7));
    
    const friday = new Date(monday);
    friday.setDate(monday.getDate() + 4);

    const startDay = monday.getDate();
    const endDay = friday.getDate();
    const startMonth = monday.toLocaleDateString('en-GB', { month: 'long' });
    const endMonth = friday.toLocaleDateString('en-GB', { month: 'long' });
    const year = friday.getFullYear();

    if (startMonth === endMonth) {
      return `${startDay}–${endDay} ${startMonth} ${year}`;
    } else {
      return `${startDay} ${startMonth} – ${endDay} ${endMonth} ${year}`;
    }
  });
};

const getAllLeaveDates = (startDateStr: string, annualLeave: boolean[][]) => {
  const startDate = new Date(startDateStr);
  const leaveDates: Date[] = [];
  
  annualLeave.forEach((week, weekIdx) => {
      week.forEach((isOff, dayIdx) => {
          if (isOff) {
              const d = new Date(startDate);
              d.setDate(startDate.getDate() + (weekIdx * 7) + dayIdx);
              leaveDates.push(d);
          }
      });
  });
  
  if (leaveDates.length === 0) return '-';

  const ranges: { start: Date; end: Date }[] = [];
  let currentRange: { start: Date; end: Date } | null = null;

  leaveDates.forEach((date) => {
      if (!currentRange) {
          currentRange = { start: date, end: date };
      } else {
          const nextDay = new Date(currentRange.end);
          nextDay.setDate(nextDay.getDate() + 1);
          if (date.toDateString() === nextDay.toDateString()) {
              currentRange.end = date;
          } else {
              ranges.push(currentRange);
              currentRange = { start: date, end: date };
          }
      }
  });
  if (currentRange) ranges.push(currentRange);

  const resultParts: string[] = [];
  let currentMonthParts: string[] = [];
  let currentMonthLabel = '';

  ranges.forEach((range) => {
      const startMonth = range.start.toLocaleDateString('en-GB', { month: 'short' });
      const endMonth = range.end.toLocaleDateString('en-GB', { month: 'short' });
      const startDay = range.start.getDate();
      const endDay = range.end.getDate();

      if (startMonth !== endMonth) {
          if (currentMonthParts.length > 0) {
              resultParts.push(`${currentMonthParts.join(', ')} ${currentMonthLabel}`);
              currentMonthParts = [];
              currentMonthLabel = '';
          }
          resultParts.push(`${startDay} ${startMonth}–${endDay} ${endMonth}`);
      } else {
          if (startMonth !== currentMonthLabel) {
              if (currentMonthParts.length > 0) {
                  resultParts.push(`${currentMonthParts.join(', ')} ${currentMonthLabel}`);
              }
              currentMonthParts = [];
              currentMonthLabel = startMonth;
          }
          if (startDay === endDay) {
              currentMonthParts.push(`${startDay}`);
          } else {
              currentMonthParts.push(`${startDay}–${endDay}`);
          }
      }
  });

  if (currentMonthParts.length > 0) {
      resultParts.push(`${currentMonthParts.join(', ')} ${currentMonthLabel}`);
  }

  return resultParts.join(', ');
};

// --- Components ---

const Button = ({ children, onClick, variant = 'primary', className = '', icon: Icon }: any) => {
  const baseStyle = "flex items-center justify-center gap-2 px-4 py-2 rounded-md font-medium transition-all duration-200 text-sm disabled:opacity-50 disabled:cursor-not-allowed focus:outline-none focus:ring-2 focus:ring-[#133958] focus:ring-offset-1";
  const variants = {
    primary: "bg-[#133958] text-white hover:bg-[#003866] shadow-sm",
    secondary: "bg-white text-slate-700 border border-[#94adae] hover:bg-[#EAF0F0] shadow-sm",
    danger: "bg-red-50 text-red-600 hover:bg-red-100 border border-red-200",
    ghost: "text-slate-700 hover:bg-[#EAF0F0]"
  };
  
  return (
    <button onClick={onClick} className={`${baseStyle} ${variants[variant as keyof typeof variants]} ${className}`}>
      {Icon && <Icon size={16} />}
      {children}
    </button>
  );
};

const Card = ({ children, className = '' }: any) => (
  <div className={`bg-white rounded-xl shadow-sm border border-[#94adae] ${className}`}>
    {children}
  </div>
);

const MultiSelect = ({ options, value = [], onChange, placeholder, disabled }: any) => {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const toggleOption = (option: string) => {
    if (disabled) return;
    const newValue = value.includes(option)
      ? value.filter((v: string) => v !== option)
      : [...value, option];
    onChange(newValue);
  };

  return (
    <div className="relative" ref={containerRef}>
      <div
        className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white focus-within:ring-2 focus-within:ring-blue-500 focus-within:border-blue-500 flex justify-between items-center min-h-[38px] ${disabled ? 'bg-[#f5f8f8] cursor-not-allowed opacity-70' : 'cursor-pointer'}`}
        onClick={() => !disabled && setIsOpen(!isOpen)}
      >
        <div className="flex flex-wrap gap-1">
          {value.length === 0 ? (
            <span className="text-slate-500">{placeholder}</span>
          ) : (
            value.map((v: string) => (
              <span key={v} className="bg-[#f5f8f8] text-slate-700 text-xs px-2 py-0.5 rounded-full border border-slate-200">
                {v}
              </span>
            ))
          )}
        </div>
        <ChevronDown size={16} className={`text-slate-400 transition-transform ${isOpen ? 'rotate-180' : ''}`} />
      </div>

      {isOpen && !disabled && (
        <div className="absolute z-50 w-full mt-1 bg-white border border-slate-200 rounded-md shadow-lg max-h-60 overflow-auto animate-in fade-in zoom-in-95 duration-100">
          {options.map((opt: string) => {
            const isSelected = value.includes(opt);
            return (
              <div 
                key={opt} 
                className={`flex items-center px-3 py-2 cursor-pointer text-sm ${isSelected ? 'bg-blue-50 text-blue-700' : 'text-slate-700 hover:bg-[#f5f8f8]'}`}
                onClick={() => toggleOption(opt)}
              >
                <div className={`w-4 h-4 mr-2 rounded border flex items-center justify-center transition-colors ${isSelected ? 'bg-blue-600 border-blue-600' : 'border-slate-300 bg-white'}`}>
                  {isSelected && <Check size={12} className="text-white" />}
                </div>
                {opt}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

// --- Main Application ---

export default function CapacityTrackerApp() {
  const [view, setView] = useState<'login' | 'employee' | 'management' | 'operations' | 'settings'>('login');
  const [currentUser, setCurrentUser] = useState<string>('');
  const [db, setDb] = useState(MOCK_DB);

  // App Settings State (lifted from constants)
  const [offices, setOffices] = useState<string[]>(INITIAL_OFFICES);
  const [mentors, setMentors] = useState<string[]>(INITIAL_MENTORS);
  const [languages, setLanguages] = useState<string[]>(INITIAL_LANGUAGES);
  const [employees, setEmployees] = useState<Employee[]>(INITIAL_EMPLOYEES);
  
  // Login Logic
  const handleLogin = (action: LoginAction) => {
    if (action.type === 'it') {
      setCurrentUser('IT');
      setView('settings');
      return;
    }

    if (action.type === 'management') {
      setCurrentUser('Admin');
      setView('management');
      return;
    }

    if (action.type === 'operations') {
      setCurrentUser('Admin');
      setView('operations');
      return;
    }

    setCurrentUser(action.employeeName);
    setView('employee');
  };

  const handleLogout = () => {
    setCurrentUser('');
    setView('login');
  };

  // Save Logic
  const handleSave = (entry: WeeklyEntry) => {
    const key = `${entry.employeeName}-${entry.weekDate}`;
    setDb(prev => ({ ...prev, [key]: entry }));
    alert("Data saved successfully!");
  };

  return (
    <div className="min-h-screen bg-[#f5f8f8] text-slate-900 font-sans selection:bg-[#94adae] selection:text-white">
      {/* Header */}
      <header className="sticky top-0 z-10 flex flex-col shadow-sm">
        <div className="bg-[#003866] py-2 flex justify-center items-center text-xs text-white/90">
          <span className="uppercase tracking-wider font-medium">
            Company
          </span>
        </div>
        <div className="bg-white border-b border-slate-200">
          <div className="max-w-7xl mx-auto px-4 h-20 flex items-center justify-between">
            <div className="flex items-center gap-4">
              <div className="bg-[#133958] p-3 rounded-xl flex items-center justify-center">
                <LayoutDashboard size={28} className="text-[#e6eaf0]" />
              </div>
              <div className="flex flex-col justify-center">
                <span className="text-[#7e9c9d] text-[13px] font-semibold tracking-widest uppercase leading-tight">
                  Company
                </span>
                <h1 className="text-[#133958] text-[26px] font-bold leading-tight">
                  Employee Capacity Tracker
                </h1>
              </div>
            </div>
            
            {view !== 'login' && (
              <div className="flex items-center gap-4">
                <span className="text-sm text-slate-500 bg-[#f5f8f8] px-3 py-1 rounded-full border border-slate-200">
                  Logged in as <span className="font-semibold text-[#133958]">{currentUser}</span>
                </span>
                <button 
                  onClick={handleLogout} 
                  className="flex items-center gap-2 px-4 py-2 rounded-md font-medium transition-all duration-200 text-sm text-[#133958] hover:bg-[#f5f8f8] border border-transparent hover:border-[#94adae]"
                >
                  <LogOut size={16} />
                  Logout
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 py-8">
        {view === 'login' && (
          <LoginScreen 
            onLogin={handleLogin} 
            employees={employees}
          />
        )}
        {view === 'employee' && (
          <EmployeeForm 
            user={currentUser} 
            initialData={db[`${currentUser}-2026-02-02`]} 
            onSave={handleSave}
            offices={offices}
            mentors={mentors}
            languages={languages}
          />
        )}
        {view === 'management' && (
          <Dashboard 
            db={db} 
            offices={offices} 
            languages={languages} 
            mentors={mentors} 
          />
        )}
        {view === 'operations' && (
          <TeamDashboard db={db} />
        )}
        {view === 'settings' && (
          <SettingsView 
            offices={offices} setOffices={setOffices}
            mentors={mentors} setMentors={setMentors}
            languages={languages} setLanguages={setLanguages}
            employees={employees} setEmployees={setEmployees}
            onBack={handleLogout}
          />
        )}
      </main>
    </div>
  );
}

// --- Settings View ---

function SettingsView({ 
  offices, setOffices, 
  mentors, setMentors, 
  languages, setLanguages, 
  employees, setEmployees,
  onBack 
}: any) {
  const [newOffice, setNewOffice] = useState('');
  const [newMentor, setNewMentor] = useState('');
  const [newLanguage, setNewLanguage] = useState('');
  const [newEmpName, setNewEmpName] = useState('');
  const [newEmpPass, setNewEmpPass] = useState('');

  const addItem = (list: string[], setList: any, item: string, setItem: any) => {
    if (item.trim() && !list.includes(item.trim())) {
      setList([...list, item.trim()]);
      setItem('');
    }
  };

  const removeItem = (list: string[], setList: any, item: string) => {
    setList(list.filter(i => i !== item));
  };

  const addEmployee = () => {
    if (newEmpName.trim() && newEmpPass.trim()) {
      setEmployees([...employees, { name: newEmpName.trim(), password: newEmpPass.trim() }]);
      setNewEmpName('');
      setNewEmpPass('');
    }
  };

  const removeEmployee = (name: string) => {
    setEmployees(employees.filter((e: Employee) => e.name !== name));
  };

  const ListManager = ({ title, items, newItem, setNewItem, onAdd, onRemove }: any) => (
    <Card className="p-6 h-full flex flex-col">
      <h3 className="font-bold text-[#133958] mb-4">{title}</h3>
      <div className="flex gap-2 mb-4">
        <input 
          className="flex-1 p-2 border border-slate-300 rounded-md text-sm"
          placeholder={`Add ${title}...`}
          value={newItem}
          onChange={e => setNewItem(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && onAdd()}
        />
        <Button onClick={onAdd} variant="secondary" className="px-3"><Plus size={16} /></Button>
      </div>
      <div className="flex-1 overflow-auto max-h-60 space-y-2">
        {items.map((item: string) => (
          <div key={item} className="flex justify-between items-center p-2 bg-[#f5f8f8] rounded border border-slate-100 text-sm">
            <span>{item}</span>
            <button onClick={() => onRemove(item)} className="text-slate-400 hover:text-red-500"><X size={14} /></button>
          </div>
        ))}
      </div>
    </Card>
  );

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-4">
        <button onClick={onBack} className="flex items-center text-sm text-[#678b8c] hover:text-[#133958] font-medium transition-colors">
          <ArrowLeft size={16} className="mr-1" /> Back
        </button>
        <h2 className="text-2xl font-bold text-[#133958]">IT Configuration</h2>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Employee Management */}
        <Card className="p-6 md:col-span-2">
          <h3 className="font-bold text-[#133958] mb-4 flex items-center gap-2"><Users size={20}/> Employee Management</h3>
          <div className="flex flex-col md:flex-row gap-2 mb-4 p-4 bg-[#f5f8f8] rounded-lg border border-[#94adae]">
            <input 
              className="flex-1 p-2 border border-slate-300 rounded-md text-sm"
              placeholder="Employee Name"
              value={newEmpName}
              onChange={e => setNewEmpName(e.target.value)}
            />
            <input 
              className="flex-1 p-2 border border-slate-300 rounded-md text-sm"
              placeholder="Set Password"
              value={newEmpPass}
              onChange={e => setNewEmpPass(e.target.value)}
            />
            <Button onClick={addEmployee}>Add Employee</Button>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm text-slate-600">
              <thead className="bg-[#f5f8f8] border-b border-[#94adae]">
                <tr>
                  <th className="px-4 py-3 font-semibold text-[#133958]">Name</th>
                  <th className="px-4 py-3 font-semibold text-[#133958]">Password</th>
                  <th className="px-4 py-3 font-semibold text-right text-[#133958]">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {employees.map((emp: Employee) => (
                  <tr key={emp.name}>
                    <td className="px-4 py-3 font-medium text-slate-900">{emp.name}</td>
                    <td className="px-4 py-3 font-mono text-xs">{emp.password}</td>
                    <td className="px-4 py-3 text-right">
                      <button onClick={() => removeEmployee(emp.name)} className="text-red-500 hover:text-red-700 font-medium text-xs">Remove</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>

        {/* Lists Management */}
        <ListManager 
          title="Offices" 
          items={offices} 
          newItem={newOffice} 
          setNewItem={setNewOffice} 
          onAdd={() => addItem(offices, setOffices, newOffice, setNewOffice)}
          onRemove={(item: string) => removeItem(offices, setOffices, item)}
        />
        <ListManager 
          title="Mentors" 
          items={mentors} 
          newItem={newMentor} 
          setNewItem={setNewMentor} 
          onAdd={() => addItem(mentors, setMentors, newMentor, setNewMentor)}
          onRemove={(item: string) => removeItem(mentors, setMentors, item)}
        />
        <ListManager 
          title="Languages" 
          items={languages} 
          newItem={newLanguage} 
          setNewItem={setNewLanguage} 
          onAdd={() => addItem(languages, setLanguages, newLanguage, setNewLanguage)}
          onRemove={(item: string) => removeItem(languages, setLanguages, item)}
        />
      </div>
    </div>
  );
}

// --- Sub-Views ---

function LoginScreen({ onLogin, employees }: { onLogin: (action: LoginAction) => void, employees: Employee[] }) {
  const [step, setStep] = useState<'select' | 'auth'>('select');
  const [role, setRole] = useState<'employee' | 'management' | 'operations' | 'it' | null>(null);
  const [selectedEmployee, setSelectedEmployee] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  const IT_PASSWORD = 'itpass123';
  const ADMIN_PASSWORD = 'admin123';

  const handleRoleSelect = (selectedRole: 'employee' | 'management' | 'operations' | 'it') => {
    setRole(selectedRole);
    setStep('auth');
    setError('');
    setPassword('');
    setSelectedEmployee(''); 
  };

  const handleBack = () => {
    setStep('select');
    setRole(null);
    setError('');
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError('');

    if (role === 'it') {
      if (password === IT_PASSWORD) onLogin({ type: 'it' });
      else setError('Invalid IT Password');
    } else if (role === 'management') {
      if (password === ADMIN_PASSWORD) onLogin({ type: 'management' });
      else setError('Invalid Manager Password');
    } else if (role === 'operations') {
      if (password === ADMIN_PASSWORD) onLogin({ type: 'operations' });
      else setError('Invalid Team Dashboard Password');
    } else {
      if (!selectedEmployee) {
        setError('Please select an Employee.');
        return;
      }
      const emp = employees.find(e => e.name === selectedEmployee);
      if (emp && password === emp.password) onLogin({ type: 'employee', employeeName: selectedEmployee });
      else setError('Invalid Password');
    }
  };

  if (step === 'select') {
    return (
      <div className="max-w-md mx-auto mt-20 animate-in fade-in slide-in-from-bottom-4 duration-500">
        <Card className="p-8 space-y-6 text-center relative">
          {/* IT Settings Button */}
          <button 
            onClick={() => handleRoleSelect('it')}
            className="absolute top-4 right-4 text-[#678b8c] hover:bg-[#EAF0F0] rounded-md transition-colors p-2 focus:outline-none focus:ring-2 focus:ring-[#133958]"
            title="IT Settings"
          >
            <Settings size={20} />
          </button>

          <div className="space-y-2">
            <h2 className="text-2xl font-bold text-[#133958]">Welcome Back</h2>
            <p className="text-[#678b8c]">Select your access level to continue</p>
          </div>
          
          <div className="space-y-3">
            <button 
              onClick={() => handleRoleSelect('employee')}
              className="w-full group relative flex items-center p-4 bg-white border border-[#94adae] rounded-xl hover:bg-[#EAF0F0] transition-all text-left focus:outline-none focus:ring-2 focus:ring-[#133958]"
            >
              <div className="p-3 bg-[#EAF0F0] rounded-lg mr-4 transition-colors">
                <Users size={24} className="text-[#133958]" />
              </div>
              <div className="flex-1">
                <div className="flex items-center justify-between">
                  <p className="font-semibold text-slate-700">Employee Login</p>
                  <ChevronRight size={16} className="text-[#678b8c] opacity-0 group-hover:opacity-100 transition-opacity" />
                </div>
                <p className="text-xs text-[#678b8c]">Access your weekly capacity form</p>
              </div>
            </button>

            <button 
              onClick={() => handleRoleSelect('operations')}
              className="w-full group relative flex items-center p-4 bg-white border border-[#94adae] rounded-xl hover:bg-[#EAF0F0] transition-all text-left focus:outline-none focus:ring-2 focus:ring-[#133958]"
            >
              <div className="p-3 bg-[#EAF0F0] rounded-lg mr-4 transition-colors">
                <LayoutDashboard size={24} className="text-[#133958]" />
              </div>
              <div className="flex-1">
                <div className="flex items-center justify-between">
                  <p className="font-semibold text-slate-700">Team Dashboard</p>
                  <ChevronRight size={16} className="text-[#678b8c] opacity-0 group-hover:opacity-100 transition-opacity" />
                </div>
                <p className="text-xs text-[#678b8c]">View weekly busy and available Employees</p>
              </div>
            </button>

            <button 
              onClick={() => handleRoleSelect('management')}
              className="w-full group relative flex items-center p-4 bg-white border border-[#94adae] rounded-xl hover:bg-[#EAF0F0] transition-all text-left focus:outline-none focus:ring-2 focus:ring-[#133958]"
            >
              <div className="p-3 bg-[#EAF0F0] rounded-lg mr-4 transition-colors">
                <LayoutDashboard size={24} className="text-[#133958]" />
              </div>
              <div className="flex-1">
                <div className="flex items-center justify-between">
                  <p className="font-semibold text-slate-700">Management Dashboard</p>
                  <ChevronRight size={16} className="text-[#678b8c] opacity-0 group-hover:opacity-100 transition-opacity" />
                </div>
                <p className="text-xs text-[#678b8c]">View full Employee management overview</p>
              </div>
            </button>
          </div>
          
          <p className="text-xs text-[#94adae] mt-6">Please contact IT if you have any questions</p>
        </Card>
      </div>
    );
  }

  return (
    <div className="max-w-sm mx-auto mt-20 animate-in fade-in slide-in-from-right-8 duration-300">
      <Card className="p-8">
        <button 
          onClick={handleBack}
          className="flex items-center text-xs text-[#678b8c] hover:bg-[#EAF0F0] px-2 py-1 -ml-2 rounded-md transition-colors mb-6 group focus:outline-none focus:ring-2 focus:ring-[#133958]"
        >
          <ArrowLeft size={14} className="mr-1 group-hover:-translate-x-1 transition-transform" /> 
          Back to selection
        </button>

        <div className="text-center mb-8">
          <div className={`mx-auto w-12 h-12 rounded-full flex items-center justify-center mb-3 bg-[#EAF0F0]`}>
            {role === 'it' ? <Settings size={24} className="text-[#133958]" /> : <Lock size={24} className="text-[#133958]" />}
          </div>
          <h2 className="text-xl font-bold text-[#133958]">
            {role === 'management'
              ? 'Manager Access'
              : role === 'operations'
                ? 'Team Dashboard Access'
                : role === 'it'
                  ? 'IT Configuration'
                  : 'Employee Access'}
          </h2>
          <p className="text-sm text-[#678b8c]">
            {role === 'management' || role === 'operations'
              ? 'Enter admin password'
              : role === 'it'
                ? 'Enter IT master password'
                : 'Log in to view your data'}
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          {role === 'employee' && (
            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-[#678b8c] uppercase tracking-wider">Employee Name</label>
              <select 
                className="w-full p-2.5 bg-white border border-[#94adae] rounded-lg text-sm focus:ring-2 focus:ring-[#133958] outline-none transition-all"
                value={selectedEmployee}
                onChange={(e) => setSelectedEmployee(e.target.value)}
              >
                <option value="" disabled>Select an Employee...</option>
                {employees.map(emp => (
                  <option key={emp.name} value={emp.name}>{emp.name}</option>
                ))}
              </select>
            </div>
          )}

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-[#678b8c] uppercase tracking-wider">Password</label>
            <input 
              type="password"
              placeholder="••••••••"
              autoFocus
              className="w-full p-2.5 bg-white border border-[#94adae] rounded-lg text-sm focus:ring-2 focus:ring-[#133958] outline-none transition-all"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </div>

          {error && (
            <div className="flex items-center gap-2 text-xs text-[#003866] bg-[#94adae]/20 p-3 rounded-lg border border-[#94adae]">
              <AlertCircle size={14} className="text-[#133958]" />
              {error}
            </div>
          )}

          <Button className="w-full justify-center mt-2" onClick={() => {}}>
            Login
          </Button>

          <div className="text-center mt-4">
            <p className="text-xs text-[#94adae]">
              {role === 'it'
                ? 'Default: itpass123'
                : role === 'management' || role === 'operations'
                  ? 'Default: admin123'
                  : 'Default: pass123'}
            </p>
          </div>
        </form>
      </Card>
    </div>
  );
}

function EmployeeForm({ user, initialData, onSave, readOnly = false, onBack, offices, mentors, languages }: any) {
  // Default structure if no data exists
  const defaultData: WeeklyEntry = {
    weekDate: '2026-02-02',
    employeeName: user,
    office: offices[0],
    mentor: '',
    languages: ['English'],
    annualLeave: Array(4).fill(null).map(() => Array(5).fill(false)),
    selfAssessment: 'Open',
    lastUpdated: new Date().toISOString(),
    projects: []
  };

  const [formData, setFormData] = useState<WeeklyEntry>(initialData || defaultData);
  const weekLabels = getWeekLabels(formData.weekDate);

  // Calculate totals for all 4 weeks
  // Includes Project Capacity + Annual Leave (1 day = 20%)
  const weeklyTotals = [0, 1, 2, 3].map(i => {
    const projectSum = formData.projects.reduce((sum, p) => sum + (p.capacities[i] || 0), 0);
    // Count true values in boolean array for this week
    const leaveCount = (formData.annualLeave[i] || []).filter(Boolean).length;
    const leavePct = (leaveCount / 5) * 100;
    return Math.round(projectSum + leavePct);
  });
  
  // Status Color Logic (Updated to new thresholds)
  const getStatusColor = (pct: number) => {
    if (pct >= 80) return 'text-red-600 bg-red-50 border-red-200';
    if (pct >= 60) return 'text-orange-600 bg-orange-50 border-orange-200';
    if (pct >= 40) return 'text-blue-600 bg-blue-50 border-blue-200';
    return 'text-green-600 bg-green-50 border-green-200';
  };

  const addProject = () => {
    const newProject: Project = {
      id: Math.random().toString(36).substr(2, 9),
      name: '',
      category: 'Category A',
      owner: '',
      capacities: [0, 0, 0, 0],
      deadline: ''
    };
    setFormData({ ...formData, projects: [...formData.projects, newProject] });
  };

  const removeProject = (id: string) => {
    setFormData({ ...formData, projects: formData.projects.filter(p => p.id !== id) });
  };

  const updateProject = (id: string, field: keyof Project, value: any) => {
    setFormData({
      ...formData,
      projects: formData.projects.map(p => p.id === id ? { ...p, [field]: value } : p)
    });
  };

  const updateCapacity = (id: string, index: number, value: number) => {
    setFormData({
      ...formData,
      projects: formData.projects.map(p => {
        if (p.id !== id) return p;
        const newCapacities = [...p.capacities] as [number, number, number, number];
        newCapacities[index] = value;
        return { ...p, capacities: newCapacities };
      })
    });
  };

  const toggleLeaveDay = (weekIndex: number, dayIndex: number) => {
    if (readOnly) return;
    const newLeave = formData.annualLeave.map((week, wIdx) => {
      if (wIdx === weekIndex) {
        const newDays = [...week];
        newDays[dayIndex] = !newDays[dayIndex];
        return newDays;
      }
      return week;
    });
    setFormData({ ...formData, annualLeave: newLeave });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Header and Controls */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-end">
        
        {/* Title Section - Aligned with Profile */}
        <div className="lg:col-span-1">
          {readOnly && onBack && (
            <button 
              onClick={onBack}
              className="flex items-center text-xs text-[#678b8c] hover:text-[#133958] mb-4 group transition-colors"
            >
              <ArrowLeft size={14} className="mr-1 group-hover:-translate-x-1 transition-transform" /> 
              Back to Dashboard
            </button>
          )}
          <h2 className="text-2xl font-bold text-[#133958]">
            {readOnly ? 'Employee Profile' : 'Monthly Capacity Form'}
          </h2>
          <p className="text-[#678b8c]">Week commencing {formData.weekDate}</p>
        </div>
        
        {/* Stats & Actions - Aligned with Active Projects */}
        <div className="lg:col-span-2 flex flex-col md:flex-row items-end md:items-center justify-between gap-4">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 w-full">
                {weeklyTotals.map((total, idx) => {
                    const label = weekLabels[idx].replace(/ \d{4}$/, ''); // Remove year for compactness
                    const colorClass = getStatusColor(total);
                    
                    return (
                        <div key={idx} className={`flex flex-col items-center justify-center px-4 py-2 rounded-lg border ${colorClass}`}>
                            <span className="text-[10px] font-bold uppercase tracking-tight mb-0.5 opacity-80 text-center">{label}</span>
                            <div className="flex items-center gap-1.5">
                                <span className="text-2xl font-bold leading-none">{total}%</span>
                                {total > 100 && <AlertCircle size={16} />}
                            </div>
                        </div>
                    );
                })}
            </div>
            {!readOnly && (
              <Button onClick={() => onSave(formData)} icon={Save} className="w-full md:w-auto py-4 md:py-2 whitespace-nowrap">Save Entry</Button>
            )}
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Sidebar Column */}
        <div className="space-y-6 h-fit lg:col-span-1">
          
          {/* Profile Card */}
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold text-[#133958] flex items-center gap-2">
              <FileText size={18} className="text-[#678b8c]" /> Profile
            </h3>
            
            <div className="space-y-3">
              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">Employee Name</label>
                <input 
                  type="text" 
                  className="w-full p-2 border border-slate-300 rounded-md text-sm bg-[#f5f8f8] text-slate-500 cursor-not-allowed"
                  value={formData.employeeName}
                  readOnly
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">Office</label>
                <select 
                  className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none ${readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''}`}
                  value={formData.office}
                  onChange={e => setFormData({...formData, office: e.target.value})}
                  disabled={readOnly}
                >
                  {offices.map((opt: string) => <option key={opt}>{opt}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">Mentor</label>
                <select 
                  className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none ${readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''}`}
                  value={formData.mentor}
                  onChange={e => setFormData({...formData, mentor: e.target.value})}
                  disabled={readOnly}
                >
                  <option value="">Select Mentor</option>
                  {mentors.map((opt: string) => <option key={opt}>{opt}</option>)}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">Languages</label>
                <MultiSelect 
                  options={languages}
                  value={formData.languages}
                  onChange={(val: string[]) => setFormData({...formData, languages: val})}
                  placeholder="Select Languages"
                  disabled={readOnly}
                />
              </div>
              
              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">Self Assessment</label>
                <select 
                  className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white ${readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''}`}
                  value={formData.selfAssessment}
                  onChange={e => setFormData({...formData, selfAssessment: e.target.value as any})}
                  disabled={readOnly}
                >
                  <option>Open</option>
                  <option>Limited Capacity</option>
                  <option>At Capacity</option>
                  <option>Over Capacity</option>
                </select>
              </div>
            </div>
          </Card>

          {/* Annual Leave Card */}
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold text-[#133958] flex items-center gap-2">
              <CalendarDays size={18} className="text-[#678b8c]" /> Annual Leave
            </h3>
            <p className="text-xs text-[#678b8c] leading-snug">
              Tick days on leave. Each day counts as 20% capacity.
            </p>
            
            <div className="space-y-4">
              {weekLabels.map((dateStr, weekIdx) => {
                // Ensure array exists for this week
                const days = formData.annualLeave[weekIdx] || [false, false, false, false, false];
                
                return (
                  <div key={weekIdx} className="bg-[#f5f8f8] p-3 rounded-md border border-slate-200">
                    <div className="text-xs font-bold text-slate-700 mb-2 text-center">
                      Week of {dateStr.replace(/ \d{4}$/, '')}
                    </div>
                    <div className="flex justify-between">
                      {WEEKDAYS.map((dayName, dayIdx) => (
                        <label key={dayIdx} className={`flex flex-col items-center gap-1 group ${readOnly ? 'cursor-not-allowed opacity-70' : 'cursor-pointer'}`}>
                          <span className="text-[10px] text-slate-500 font-medium group-hover:text-blue-600">{dayName}</span>
                          <input 
                            type="checkbox" 
                            className={`w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500 ${readOnly ? 'cursor-not-allowed' : 'cursor-pointer'}`}
                            checked={days[dayIdx]}
                            onChange={() => toggleLeaveDay(weekIdx, dayIdx)}
                            disabled={readOnly}
                          />
                        </label>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>

        </div>

        {/* Projects List */}
        <Card className="p-6 lg:col-span-2">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <h3 className="font-semibold text-[#133958]">Active Projects</h3>
              <span className="bg-[#f5f8f8] text-slate-600 px-2.5 py-0.5 rounded-full text-xs font-semibold border border-slate-200">
                {formData.projects.length}
              </span>
            </div>
            {!readOnly && (
              <Button variant="secondary" onClick={addProject} icon={Plus}>Add Project</Button>
            )}
          </div>

          <div className="space-y-6">
            {formData.projects.length === 0 && (
              <div className="text-center py-10 text-slate-400 border-2 border-dashed border-slate-200 rounded-lg bg-[#f5f8f8]">
                No projects added yet.
              </div>
            )}
            
            {formData.projects.map((project) => (
              <div key={project.id} className="p-4 bg-[#f5f8f8] rounded-lg border border-[#94adae] group hover:border-[#678b8c] transition-colors space-y-4">
                
                {/* Row 1: Basic Info */}
                <div className="grid grid-cols-12 gap-3 items-end">
                  <div className="col-span-4">
                    <label className="text-xs font-medium text-[#678b8c] ml-1">Project Name</label>
                    <input 
                      type="text" 
                      placeholder="e.g. 2019-40"
                      className={`w-full p-2 border border-slate-300 rounded-md text-sm ${readOnly ? 'bg-slate-50 text-slate-600' : ''}`}
                      value={project.name}
                      onChange={e => updateProject(project.id, 'name', e.target.value)}
                      readOnly={readOnly}
                    />
                  </div>
                  <div className="col-span-3">
                    <label className="text-xs font-medium text-[#678b8c] ml-1">Category</label>
                    <select
                      className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none ${readOnly ? 'bg-slate-50 text-slate-600 cursor-not-allowed' : ''}`}
                      value={project.category}
                      onChange={e => updateProject(project.id, 'category', e.target.value)}
                      disabled={readOnly}
                    >
                      {PROJECT_CATEGORIES.map(cat => (
                        <option key={cat} value={cat}>{cat}</option>
                      ))}
                    </select>
                  </div>
                  <div className="col-span-4">
                    <label className="text-xs font-medium text-[#678b8c] ml-1">Owner</label>
                    <select
                      className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none ${readOnly ? 'bg-slate-50 text-slate-600 cursor-not-allowed' : ''}`}
                      value={project.owner}
                      onChange={e => updateProject(project.id, 'owner', e.target.value)}
                      disabled={readOnly}
                    >
                      <option value="">Select Owner</option>
                      {project.owner && !mentors.includes(project.owner) && (
                        <option value={project.owner}>{project.owner}</option>
                      )}
                      {mentors.map((opt: string) => (
                        <option key={opt} value={opt}>{opt}</option>
                      ))}
                    </select>
                  </div>
                  <div className="col-span-1 flex justify-center pb-2">
                    {!readOnly && (
                      <button 
                        onClick={() => removeProject(project.id)}
                        className="text-slate-400 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={18} />
                      </button>
                    )}
                  </div>
                </div>

                {/* Row 2: Capacity Forecast (Updated Container) */}
                <div className="relative border border-slate-300 rounded-md pt-5 pb-3 px-3 mt-2 bg-white/50">
                  <span className="absolute -top-2.5 left-2 bg-[#f5f8f8] px-2 text-xs font-bold text-[#678b8c]">
                    Upcoming Capacity
                  </span>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
                    {weekLabels.map((dateStr, idx) => (
                      <div key={idx}>
                         <label className={`text-[10px] uppercase font-bold mb-1.5 block truncate text-center ${idx === 0 ? 'text-[#133958]' : 'text-slate-400'}`} title={dateStr}>
                           {dateStr}
                         </label>
                         <div className="relative">
                          <input 
                            type="number" 
                            min="0"
                            max="100"
                            className={`w-full p-1.5 border rounded-md text-sm text-center ${idx === 0 ? 'border-blue-300 bg-blue-50 text-blue-900 font-bold' : 'border-slate-200 text-slate-600'} ${readOnly ? 'bg-slate-50 cursor-not-allowed' : ''}`}
                            value={project.capacities[idx]}
                            onChange={e => updateCapacity(project.id, idx, parseFloat(e.target.value) || 0)}
                            readOnly={readOnly}
                          />
                          <span className="absolute right-2 top-1.5 text-xs text-slate-400 pointer-events-none">%</span>
                         </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Row 3: Details */}
                <div>
                  <label className="text-xs font-medium text-[#678b8c] ml-1">Deadline / Comments</label>
                  <input 
                    type="text" 
                    placeholder="Key deadlines or notes..."
                    className={`w-full p-2 border border-slate-300 rounded-md text-sm ${readOnly ? 'bg-slate-50 text-slate-600' : ''}`}
                    value={project.deadline}
                    onChange={e => updateProject(project.id, 'deadline', e.target.value)}
                    readOnly={readOnly}
                  />
                </div>

              </div>
            ))}
          </div>
          
        </Card>
      </div>
    </div>
  );
}

function TeamDashboard({ db }: { db: Record<string, WeeklyEntry> }) {
  const entries = Object.values(db);
  const defaultWeekDate = '2026-02-02';
  const referenceWeekDate = entries[0]?.weekDate || defaultWeekDate;
  const weekLabels = getWeekLabels(referenceWeekDate);

  const spanStartDate = new Date(referenceWeekDate);
  const spanEndDate = new Date(spanStartDate);
  spanEndDate.setDate(spanStartDate.getDate() + 25);
  const dateSpanLabel = `${spanStartDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })} to ${spanEndDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}`;

  const getLoadBadgeClass = (load: number) => {
    if (load >= 100) return 'bg-red-100 text-red-700';
    if (load >= 80) return 'bg-orange-100 text-orange-700';
    if (load >= 40) return 'bg-blue-100 text-blue-700';
    return 'bg-green-100 text-green-700';
  };

  const weeklyBreakdown = React.useMemo(() => {
    return Array.from({ length: 4 }).map((_, weekIdx) => {
      const people = entries.map((entry) => {
        const projectLoad = entry.projects.reduce((acc, project) => acc + (project.capacities[weekIdx] || 0), 0);
        const leaveCount = (entry.annualLeave[weekIdx] || []).filter(Boolean).length;
        const leaveLoad = (leaveCount / WEEKDAYS.length) * 100;
        const load = Math.round(projectLoad + leaveLoad);
        return { name: entry.employeeName, load };
      });

      const busy = people
        .filter(person => person.load >= 80)
        .sort((a, b) => a.name.localeCompare(b.name));

      const available = people
        .filter(person => person.load < 80)
        .sort((a, b) => a.name.localeCompare(b.name));

      return { busy, available };
    });
  }, [entries]);

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-2xl font-bold text-[#133958]">Team Dashboard</h2>
        <p className="text-[#678b8c]">Employee capacity dashboard {dateSpanLabel}</p>
      </div>

      {entries.length === 0 ? (
        <Card className="p-8 text-center text-slate-500">
          No data available.
        </Card>
      ) : (
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-2 2xl:grid-cols-4">
          {weeklyBreakdown.map((weekData, weekIdx) => (
            <Card key={`operations-week-${weekIdx}`} className="p-4 space-y-4">
              <div>
                <h3 className="font-semibold text-[#133958]">Week {weekIdx + 1}</h3>
                <p className="text-xs text-[#678b8c]">{weekLabels[weekIdx]}</p>
              </div>

              <div>
                <p className="text-xs uppercase tracking-wide text-[#133958] font-semibold mb-2">Busy ({weekData.busy.length})</p>
                {weekData.busy.length === 0 ? (
                  <p className="text-xs text-slate-500">No busy Employees</p>
                ) : (
                  <div className="space-y-1.5">
                    {weekData.busy.map((person) => (
                      <div key={`busy-${weekIdx}-${person.name}`} className="flex items-center justify-between text-sm">
                        <span className="text-slate-700 truncate pr-2">{person.name}</span>
                        <span className={`px-2 py-1 rounded-full text-xs font-bold ${getLoadBadgeClass(person.load)}`}>
                          {person.load}%
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div>
                <p className="text-xs uppercase tracking-wide text-[#133958] font-semibold mb-2">Available ({weekData.available.length})</p>
                {weekData.available.length === 0 ? (
                  <p className="text-xs text-slate-500">No available Employees</p>
                ) : (
                  <div className="space-y-1.5">
                    {weekData.available.map((person) => (
                      <div key={`available-${weekIdx}-${person.name}`} className="flex items-center justify-between text-sm">
                        <span className="text-slate-700 truncate pr-2">{person.name}</span>
                        <span className={`px-2 py-1 rounded-full text-xs font-bold ${getLoadBadgeClass(person.load)}`}>
                          {person.load}%
                        </span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}


function Dashboard({ db, offices, languages }: { db: Record<string, WeeklyEntry>, offices: string[], languages: string[], mentors: string[] }) {
  type DashboardEntry = WeeklyEntry & {
    weeklyLoads: number[];
    averageLoad: number;
    loadDelta: number;
    weekLoad1: number;
    weekLoad2: number;
    weekLoad3: number;
    weekLoad4: number;
    categoryACount: number;
    categoryBCount: number;
  };

  const [sortConfig, setSortConfig] = useState<{ key: string; direction: 'asc' | 'desc' } | null>(null);
  const [activeWeek, setActiveWeek] = useState(0);
  const [peopleLoadMode, setPeopleLoadMode] = useState<'most' | 'least'>('most');
  const [selectedOffice, setSelectedOffice] = useState('All Offices');
  const [selectedLanguage, setSelectedLanguage] = useState('All Languages');
  const [selectedEntry, setSelectedEntry] = useState<WeeklyEntry | null>(null);

  const defaultWeekDate = '2026-02-02';
  const referenceWeekDate = Object.values(db)[0]?.weekDate || defaultWeekDate;
  const weekLabels = getWeekLabels(referenceWeekDate);
  const relativeWeekLabels = ['This week', 'Next week', 'In two weeks', 'In three weeks'];
  const spanStartDate = new Date(referenceWeekDate);
  const spanEndDate = new Date(spanStartDate);
  spanEndDate.setDate(spanStartDate.getDate() + 25);
  const dateSpanLabel = `${spanStartDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })} to ${spanEndDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}`;
  const activeWeekKey = `weekLoad${activeWeek + 1}`;

  const getLoadBadgeClass = (load: number) => {
    if (load >= 100) return 'bg-red-100 text-red-700';
    if (load >= 80) return 'bg-orange-100 text-orange-700';
    if (load >= 40) return 'bg-blue-100 text-blue-700';
    return 'bg-green-100 text-green-700';
  };

  const getProjectBadgeClass = (load: number) => {
    if (load >= 50) return 'bg-red-100 text-red-700 border border-red-200';
    if (load >= 25) return 'bg-orange-100 text-orange-700 border border-orange-200';
    if (load > 0) return 'bg-blue-50 text-blue-700 border border-blue-200';
    return 'bg-[#f5f8f8] text-[#678b8c] border border-[#94adae]';
  };

  const tableData = React.useMemo<DashboardEntry[]>(() => {
    return Object.values(db).map((entry) => {
      const weeklyLoads = Array.from({ length: 4 }).map((_, weekIdx) => {
        const projectLoad = entry.projects.reduce((acc, project) => acc + (project.capacities[weekIdx] || 0), 0);
        const leaveCount = (entry.annualLeave[weekIdx] || []).filter(Boolean).length;
        const leaveLoad = (leaveCount / WEEKDAYS.length) * 100;
        return Math.round(projectLoad + leaveLoad);
      });

      const averageLoad = Math.round(weeklyLoads.reduce((acc, load) => acc + load, 0) / weeklyLoads.length);
      const categoryACount = entry.projects.filter(project => project.category === 'Category A').length;
      const categoryBCount = entry.projects.filter(project => project.category === 'Category B').length;

      return {
        ...entry,
        weeklyLoads,
        averageLoad,
        loadDelta: weeklyLoads[3] - weeklyLoads[0],
        weekLoad1: weeklyLoads[0],
        weekLoad2: weeklyLoads[1],
        weekLoad3: weeklyLoads[2],
        weekLoad4: weeklyLoads[3],
        categoryACount,
        categoryBCount
      };
    });
  }, [db]);

  const filteredData = React.useMemo(() => {
    return tableData.filter((entry) => {
      const matchesOffice = selectedOffice === 'All Offices' || entry.office === selectedOffice;
      const matchesLanguage = selectedLanguage === 'All Languages' || entry.languages.includes(selectedLanguage);
      return matchesOffice && matchesLanguage;
    });
  }, [tableData, selectedOffice, selectedLanguage]);

  const weeklySummary = React.useMemo(() => {
    return Array.from({ length: 4 }).map((_, weekIdx) => {
      const weekLoads = filteredData.map(entry => entry.weeklyLoads[weekIdx] || 0);
      const totalLoad = weekLoads.reduce((acc, load) => acc + load, 0);
      const avgLoad = weekLoads.length > 0 ? Math.round(totalLoad / weekLoads.length) : 0;
      const overCap = weekLoads.filter(load => load >= 100).length;
      const atCapacity = weekLoads.filter(load => load >= 80 && load < 100).length;
      const totalLeaveDays = filteredData.reduce((acc, entry) => acc + (entry.annualLeave[weekIdx] || []).filter(Boolean).length, 0);
      const avgLeaveDays = weekLoads.length > 0 ? totalLeaveDays / weekLoads.length : 0;

      return {
        avgLoad,
        overCap,
        atCapacity,
        avgLeaveDays: avgLeaveDays.toFixed(1)
      };
    });
  }, [filteredData]);

  const activeWeekInsights = React.useMemo(() => {
    const weekLoads = filteredData.map(entry => entry.weeklyLoads[activeWeek] || 0);
    const averageLoad = weekLoads.length > 0
      ? Math.round(weekLoads.reduce((acc, load) => acc + load, 0) / weekLoads.length)
      : 0;
    const atCapacityCount = weekLoads.filter(load => load >= 80 && load < 100).length;
    const overCapCount = weekLoads.filter(load => load >= 100).length;
    const lookingForWorkCount = weekLoads.filter(load => load < 80).length;

    const mostLoadedPeople = [...filteredData]
      .sort((a, b) => (b.weeklyLoads[activeWeek] || 0) - (a.weeklyLoads[activeWeek] || 0))
      .slice(0, 3)
      .map(entry => ({
        name: entry.employeeName,
        load: entry.weeklyLoads[activeWeek] || 0
      }));

    const leastLoadedPeople = [...filteredData]
      .sort((a, b) => (a.weeklyLoads[activeWeek] || 0) - (b.weeklyLoads[activeWeek] || 0))
      .slice(0, 3)
      .map(entry => ({
        name: entry.employeeName,
        load: entry.weeklyLoads[activeWeek] || 0
      }));

    const projectTotals = new Map<string, number>();
    filteredData.forEach((entry) => {
      entry.projects.forEach((project) => {
        const load = project.capacities[activeWeek] || 0;
        if (load <= 0) return;
        projectTotals.set(project.name, (projectTotals.get(project.name) || 0) + load);
      });
    });

    const topProjects = Array.from(projectTotals.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([name, total]) => ({ name, total: Math.round(total) }));

    return {
      averageLoad,
      atCapacityCount,
      overCapCount,
      lookingForWorkCount,
      mostLoadedPeople,
      leastLoadedPeople,
      topProjects
    };
  }, [filteredData, activeWeek]);

  const sortedData = React.useMemo(() => {
    const appliedSort = sortConfig || { key: activeWeekKey, direction: 'desc' as const };

    return [...filteredData].sort((a, b) => {
      const aValue: any = a[appliedSort.key as keyof DashboardEntry];
      const bValue: any = b[appliedSort.key as keyof DashboardEntry];

      if (typeof aValue === 'string' && typeof bValue === 'string') {
        const compareResult = aValue.localeCompare(bValue);
        return appliedSort.direction === 'asc' ? compareResult : -compareResult;
      }

      if (aValue < bValue) return appliedSort.direction === 'asc' ? -1 : 1;
      if (aValue > bValue) return appliedSort.direction === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredData, sortConfig, activeWeekKey]);

  const requestSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  const getSortIcon = (name: string) => {
    if (!sortConfig || sortConfig.key !== name) return <div className="w-4" />;
    return sortConfig.direction === 'asc'
      ? <ArrowUp size={14} className="text-slate-500" />
      : <ArrowDown size={14} className="text-slate-500" />;
  };

  if (selectedEntry) {
    return (
      <EmployeeForm
        user={selectedEntry.employeeName}
        initialData={selectedEntry}
        onSave={() => {}}
        readOnly={true}
        onBack={() => setSelectedEntry(null)}
        offices={offices}
        languages={languages}
        mentors={[]}
      />
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-start justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-[#133958]">Employee Capacity Overview</h2>
          <p className="text-[#678b8c]">Employee capacity dashboard {dateSpanLabel}</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <select
            className="p-2 border border-slate-300 rounded-md text-sm bg-white"
            value={selectedOffice}
            onChange={(e) => setSelectedOffice(e.target.value)}
          >
            <option>All Offices</option>
            {offices.map((office: string) => (
              <option key={office}>{office}</option>
            ))}
          </select>
          <select
            className="p-2 border border-slate-300 rounded-md text-sm bg-white"
            value={selectedLanguage}
            onChange={(e) => setSelectedLanguage(e.target.value)}
          >
            <option>All Languages</option>
            {languages.map((lang: string) => (
              <option key={lang}>{lang}</option>
            ))}
          </select>
          <Button variant="secondary">Export CSV</Button>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-3 md:grid-cols-2 xl:grid-cols-4">
        {weeklySummary.map((summary, weekIdx) => {
          const isActive = activeWeek === weekIdx;
          const shortLabel = weekLabels[weekIdx]?.replace(/ \d{4}$/, '');

          return (
            <button
              key={`week-summary-${weekIdx}`}
              type="button"
              onClick={() => setActiveWeek(weekIdx)}
              className={`text-left rounded-xl border p-4 transition-colors ${isActive ? 'border-[#133958] bg-[#e6eff3]' : 'border-[#94adae] bg-white hover:bg-[#f5f8f8]'}`}
            >
              <div className="flex items-center justify-between">
                <p className="text-sm font-semibold text-[#133958]">{relativeWeekLabels[weekIdx] || `Week ${weekIdx + 1}`}</p>
                <span className={`px-2 py-1 rounded-full text-xs font-bold ${getLoadBadgeClass(summary.avgLoad)}`}>
                  {summary.avgLoad}%
                </span>
              </div>
              <p className="mt-1 text-xs text-[#678b8c]">{shortLabel}</p>
              <p className="mt-3 text-xs text-slate-600">{summary.atCapacity} at capacity | {summary.overCap} over capacity</p>
              <p className="mt-1 text-xs text-slate-500">Avg leave: {summary.avgLeaveDays} days</p>
            </button>
          );
        })}
      </div>

      <div className="grid grid-cols-1 gap-4 xl:grid-cols-3">
        <Card className="p-4">
          <p className="text-sm font-semibold text-[#133958]">Active snapshot (Week {activeWeek + 1})</p>
          <p className="mt-2 text-3xl font-bold text-[#133958]">{activeWeekInsights.averageLoad}%</p>
          <p className="text-xs text-slate-500">Average planned load</p>
          <div className="mt-3 space-y-1 text-sm text-slate-600">
            <p>{activeWeekInsights.atCapacityCount} at capacity (80-99%)</p>
            <p>{activeWeekInsights.overCapCount} over capacity (over 100%)</p>
            <p>{activeWeekInsights.lookingForWorkCount} looking for work (below 80%)</p>
          </div>
        </Card>

        <Card className="p-4">
          <div className="flex items-center justify-between gap-2">
            <p className="text-sm font-semibold text-[#133958]">
              {peopleLoadMode === 'most' ? 'Busiest Workload' : 'Most Available'} (Week {activeWeek + 1})
            </p>
            <div className="inline-flex rounded-md border border-slate-200 bg-white p-0.5">
              <button
                type="button"
                onClick={() => setPeopleLoadMode('most')}
                className={`px-2 py-1 text-xs rounded ${peopleLoadMode === 'most' ? 'bg-[#133958] text-white' : 'text-slate-600 hover:bg-slate-100'}`}
              >
                Most
              </button>
              <button
                type="button"
                onClick={() => setPeopleLoadMode('least')}
                className={`px-2 py-1 text-xs rounded ${peopleLoadMode === 'least' ? 'bg-[#133958] text-white' : 'text-slate-600 hover:bg-slate-100'}`}
              >
                Least
              </button>
            </div>
          </div>
          <div className="mt-3 space-y-2">
            {(peopleLoadMode === 'most' ? activeWeekInsights.mostLoadedPeople : activeWeekInsights.leastLoadedPeople).length === 0 && (
              <p className="text-sm text-slate-500">No team members in the current filter.</p>
            )}
            {(peopleLoadMode === 'most' ? activeWeekInsights.mostLoadedPeople : activeWeekInsights.leastLoadedPeople).map((person) => (
              <div key={person.name} className="flex items-center justify-between text-sm">
                <span className="text-slate-700">{person.name}</span>
                <span className={`px-2 py-1 rounded-full text-xs font-bold ${getLoadBadgeClass(person.load)}`}>{person.load}%</span>
              </div>
            ))}
          </div>
        </Card>

        <Card className="p-4">
          <p className="text-sm font-semibold text-[#133958]">Top projects by demand (Week {activeWeek + 1})</p>
          <div className="mt-3 space-y-2">
            {activeWeekInsights.topProjects.length === 0 && (
              <p className="text-sm text-slate-500">No scheduled project load in this week.</p>
            )}
            {activeWeekInsights.topProjects.map((project) => (
              <div key={project.name} className="flex items-center justify-between text-sm">
                <span className="text-slate-700 truncate pr-4">{project.name}</span>
                <span className={`px-2 py-1 rounded-full text-xs font-bold ${getLoadBadgeClass(project.total)}`}>{project.total}%</span>
              </div>
            ))}
          </div>
        </Card>
      </div>

      <Card className="overflow-hidden">
        <div className="pb-2">
          <table className="w-full table-fixed text-sm text-slate-600">
            <thead className="bg-[#f5f8f8] border-b border-[#94adae] uppercase text-xs font-semibold text-[#133958] text-center">
              <tr>
                <th
                  className="px-3 py-2 cursor-pointer hover:bg-slate-100 transition-colors select-none group"
                  onClick={() => requestSort('employeeName')}
                >
                  <div className="flex items-center justify-center gap-1">
                    Employee {getSortIcon('employeeName')}
                  </div>
                </th>
                <th
                  className="px-3 py-2 cursor-pointer hover:bg-slate-100 transition-colors select-none group"
                  onClick={() => requestSort('selfAssessment')}
                >
                  <div className="flex items-center justify-center gap-1">
                    Status {getSortIcon('selfAssessment')}
                  </div>
                </th>
                {weekLabels.map((_, weekIdx) => (
                  <th
                    key={`week-col-${weekIdx}`}
                    className="px-3 py-2 cursor-pointer hover:bg-slate-100 transition-colors select-none group"
                    onClick={() => requestSort(`weekLoad${weekIdx + 1}`)}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Week {weekIdx + 1} {getSortIcon(`weekLoad${weekIdx + 1}`)}
                    </div>
                  </th>
                ))}
                <th
                  className="px-3 py-2 cursor-pointer hover:bg-slate-100 transition-colors select-none group"
                  onClick={() => requestSort('categoryACount')}
                >
                  <div className="flex items-center justify-center gap-1">
                    Category A {getSortIcon('categoryACount')}
                  </div>
                </th>
                <th
                  className="px-3 py-2 cursor-pointer hover:bg-slate-100 transition-colors select-none group"
                  onClick={() => requestSort('categoryBCount')}
                >
                  <div className="flex items-center justify-center gap-1">
                    Category B {getSortIcon('categoryBCount')}
                  </div>
                </th>
                <th className="px-3 py-2 w-[14%]">Annual Leave</th>
                <th className="px-3 py-2 w-[22%]">Top Projects (Week {activeWeek + 1})</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {sortedData.map((entry) => {
                const topProjects = [...entry.projects]
                  .filter(project => (project.capacities[activeWeek] || 0) > 0)
                  .sort((a, b) => (b.capacities[activeWeek] || 0) - (a.capacities[activeWeek] || 0))
                  .slice(0, 4);

                return (
                  <tr
                    key={`${entry.employeeName}-${entry.weekDate}`}
                    className="hover:bg-slate-50 transition-colors cursor-pointer"
                    onClick={() => setSelectedEntry(entry)}
                  >
                    <td className="px-3 py-3 font-medium text-slate-900 text-center group">
                      <span className="group-hover:text-blue-600 transition-colors">
                        {entry.employeeName}
                      </span>
                    </td>
                    <td className="px-3 py-3 text-center">{entry.selfAssessment}</td>
                    {entry.weeklyLoads.map((load, weekIdx) => (
                      <td key={`load-${entry.employeeName}-${weekIdx}`} className="px-3 py-3 text-center">
                        <span className={`px-2 py-1 rounded-full text-xs font-bold ${getLoadBadgeClass(load)}`}>
                          {load}%
                        </span>
                      </td>
                    ))}
                    <td className="px-3 py-3 text-center text-slate-600 font-medium">
                      {entry.categoryACount}
                    </td>
                    <td className="px-3 py-3 text-center text-slate-600 font-medium">
                      {entry.categoryBCount}
                    </td>
                    <td className="px-3 py-3 text-xs text-slate-600 text-center break-words">
                      {getAllLeaveDates(entry.weekDate, entry.annualLeave)}
                    </td>
                    <td className="px-3 py-3 align-top">
                      {topProjects.length === 0 ? (
                        <p className="text-xs text-slate-500 text-center">No active projects this week</p>
                      ) : (
                        <div className="flex flex-col gap-2">
                          {topProjects.map(project => {
                            const cap = project.capacities[activeWeek] || 0;

                            return (
                              <div key={project.id} className="flex flex-col text-xs w-full p-1.5 bg-[#f5f8f8] rounded border border-[#94adae]">
                                <div className="flex justify-between items-center mb-0.5">
                                  <span className="font-medium text-[#133958] truncate w-[70%]" title={project.name}>{project.name}</span>
                                  <span className={`px-1.5 py-px rounded text-[10px] font-bold ${getProjectBadgeClass(cap)}`}>
                                    {cap.toFixed(0)}%
                                  </span>
                                </div>
                                {project.deadline && (
                                  <div className="text-[#678b8c] truncate text-[10px] italic" title={project.deadline}>
                                    {project.deadline}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          </div>

      </Card>
    </div>
  );
}
