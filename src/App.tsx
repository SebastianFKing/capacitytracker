import React, { useEffect, useMemo, useRef, useState } from 'react';
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
  Eye,
  EyeOff
} from 'lucide-react';

// --- Types ---

type MatterType = 'Category1' | 'Category2' | 'Project';

type Project = {
  id: string;
  name: string;
  category: MatterType;
  matterType?: MatterType; // legacy field kept for backward compatibility
  owner: string;
  tasks?: string;
  capacities: [number, number, number, number]; // 4 weeks
};

type Availability2Weeks = 'With Capacity' | 'Limited Capacity' | 'No Capacity' | 'Over Capacity';

type WeeklyEntry = {
  weekDate: string; // week commencing (Monday)
  employeeName: string;
  office: string;
  mentor: string;
  languages: string[];
  interests: string;
  annualLeave: boolean[][]; // 4 weeks x 5 days
  availability2Weeks: Availability2Weeks;
  capacityComments: string[];
  projects: Project[]; // matters
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

// --- Constants ---

const INITIAL_OFFICES = ['Office A', 'Office B', 'Office C', 'Office D', 'Office E', 'Office F'];
const INITIAL_MENTORS = ['Mentor 1', 'Mentor 2', 'Mentor 3', 'Mentor 4'];
const INITIAL_LANGUAGES = ['English', 'French', 'German', 'Dutch', 'Spanish', 'Mandarin', 'Arabic'];
const PROJECT_CATEGORIES = ['Category1', 'Category2', 'Project'] as const;
const WEEKDAYS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
const HOURS_PER_WEEK = 40;
const LEAVE_HOURS_PER_DAY = HOURS_PER_WEEK / WEEKDAYS.length;
const COMMENT_WORD_LIMIT = 250;
const IT_MASTER_PASSWORD = 'itpass123';
const STORAGE_KEY_DB = 'capacity_tracker_db_v1';

const INITIAL_EMPLOYEES: Employee[] = [
  { name: 'Employee A', password: 'pass123' },
  { name: 'Employee B', password: 'pass123' },
  { name: 'Employee C', password: 'pass123' },
  { name: 'Employee D', password: 'pass123' },
  { name: 'Employee E', password: 'pass123' },
  { name: 'Employee F', password: 'pass123' }
];

// --- Mock Data ---

const baseLeave = Array(4)
  .fill(null)
  .map(() => Array(5).fill(false));

const MOCK_DB: Record<string, WeeklyEntry> = {
  'Employee A-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee A',
    office: 'Office A',
    mentor: 'Mentor 2',
    languages: ['English', 'Spanish'],
    interests: '',
    annualLeave: baseLeave,
    availability2Weeks: 'Limited Capacity',
    capacityComments: Array(4).fill(''),
    lastUpdated: new Date().toISOString(),
    projects: [
      {
        id: '1',
        name: 'Task1',
        category: 'Category1',
        matterType: 'Category1',
        owner: 'Supervisor 1',
        tasks: '',
        capacities: [25, 25, 20, 10]
      },
      {
        id: '2',
        name: 'Task2',
        category: 'Category1',
        matterType: 'Category1',
        owner: 'Supervisor 2',
        tasks: '',
        capacities: [20, 20, 15, 10]
      }
    ]
  },
  'Employee B-2026-02-02': {
    weekDate: '2026-02-02',
    employeeName: 'Employee B',
    office: 'Office E',
    mentor: 'Mentor 1',
    languages: ['French'],
    interests: '',
    annualLeave: [
      [true, false, false, false, false],
      [false, false, false, false, false],
      [false, false, false, false, false],
      [true, true, true, true, true]
    ],
    availability2Weeks: 'No Capacity',
    capacityComments: Array(4).fill(''),
    lastUpdated: new Date().toISOString(),
    projects: []
  }
};

// --- Helper Functions ---

const parseIsoDateLocal = (dateStr: string) => {
  const [yearStr, monthStr, dayStr] = dateStr.split('-');
  const year = Number(yearStr);
  const month = Number(monthStr);
  const day = Number(dayStr);

  if (
    Number.isInteger(year) &&
    Number.isInteger(month) &&
    Number.isInteger(day) &&
    month >= 1 &&
    month <= 12 &&
    day >= 1 &&
    day <= 31
  ) {
    return new Date(year, month - 1, day);
  }
  return new Date(dateStr);
};

const formatIsoDateLocal = (date: Date) => {
  const y = date.getFullYear();
  const m = `${date.getMonth() + 1}`.padStart(2, '0');
  const d = `${date.getDate()}`.padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const getCurrentWeekStart = () => {
  const now = new Date();
  const day = now.getDay();
  const diff = (day + 6) % 7; // shift so Monday is start of week
  const monday = new Date(now);
  monday.setHours(0, 0, 0, 0);
  monday.setDate(now.getDate() - diff);
  return formatIsoDateLocal(monday);
};

const formatDisplayDate = (dateStr: string) =>
  parseIsoDateLocal(dateStr).toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });

const clampCapacity = (value: number) => {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.round(value * 1000) / 1000);
};

const normalizeCapacityInput = (value: unknown) => {
  if (typeof value === 'number') return clampCapacity(value);
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return 0;
    const digitsOnly = trimmed.replace(/\D+/g, '');
    if (digitsOnly === '') return 0;
    const withoutLeadingZeros = digitsOnly.replace(/^0+(?=\d)/, '');
    return clampCapacity(Number(withoutLeadingZeros));
  }
  return clampCapacity(Number(value));
};

const percentToHours = (percent: number) => {
  if (!Number.isFinite(percent)) return 0;
  return Math.max(0, (percent / 100) * HOURS_PER_WEEK);
};

const hoursToPercent = (hours: number) => {
  if (!Number.isFinite(hours)) return 0;
  return clampCapacity((hours / HOURS_PER_WEEK) * 100);
};

const normalizeHoursInput = (value: unknown) => {
  if (typeof value === 'number') return Math.max(0, value);
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return 0;
    const compact = trimmed.replace(/\s+/g, '');
    if (compact.includes(':')) {
      const [rawHours = '', rawMinutes = ''] = compact.split(':', 2);
      const hoursDigits = rawHours.replace(/\D+/g, '');
      const minuteDigits = rawMinutes.replace(/\D+/g, '');
      if (hoursDigits === '' && minuteDigits === '') return 0;
      const hoursValue = hoursDigits === '' ? 0 : Number(hoursDigits);
      const minutesValue = minuteDigits === '' ? 0 : Number(minuteDigits);
      if (!Number.isFinite(hoursValue) || !Number.isFinite(minutesValue)) return 0;
      return Math.max(0, (hoursValue * 60 + minutesValue) / 60);
    }
    const numericValue = Number(compact);
    if (Number.isFinite(numericValue)) return Math.max(0, numericValue);
    const digitsOnly = compact.replace(/\D+/g, '');
    if (digitsOnly === '') return 0;
    return Math.max(0, Number(digitsOnly));
  }
  const fallbackValue = Number(value);
  return Number.isFinite(fallbackValue) ? Math.max(0, fallbackValue) : 0;
};

const formatHoursInput = (hours: number) => {
  if (!Number.isFinite(hours)) return '0:00';
  const totalMinutes = Math.max(0, Math.round(hours * 60));
  const wholeHours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${wholeHours}:${String(minutes).padStart(2, '0')}`;
};

const limitToWordCount = (value: string, maxWords: number) => {
  const wordsWithSpacing = value.match(/\S+\s*/g) || [];
  if (wordsWithSpacing.length <= maxWords) return value;
  return wordsWithSpacing.slice(0, maxWords).join('').trimEnd();
};

const clampWeekLoadPercent = (value: number) => {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(100, Math.round(value)));
};

const formatLeaveDaySpans = (weekLeave: boolean[]) => {
  const dayLabels = ['Mon', 'Tues', 'Wed', 'Thurs', 'Fri'];
  const selectedIndexes = weekLeave
    .map((isOff, idx) => (isOff ? idx : -1))
    .filter((idx) => idx >= 0);

  if (selectedIndexes.length === 0) return '-';

  const segments: string[] = [];
  let rangeStart = selectedIndexes[0];
  let rangeEnd = selectedIndexes[0];

  const flushRange = () => {
    if (rangeStart === rangeEnd) {
      segments.push(dayLabels[rangeStart] || WEEKDAYS[rangeStart] || `Day ${rangeStart + 1}`);
      return;
    }
    const startLabel = dayLabels[rangeStart] || WEEKDAYS[rangeStart] || `Day ${rangeStart + 1}`;
    const endLabel = dayLabels[rangeEnd] || WEEKDAYS[rangeEnd] || `Day ${rangeEnd + 1}`;
    segments.push(`${startLabel}–${endLabel}`);
  };

  for (let i = 1; i < selectedIndexes.length; i += 1) {
    const nextIdx = selectedIndexes[i];
    if (nextIdx === rangeEnd + 1) {
      rangeEnd = nextIdx;
      continue;
    }
    flushRange();
    rangeStart = nextIdx;
    rangeEnd = nextIdx;
  }
  flushRange();

  return segments.join(', ');
};

const getLatestEntryForEmployee = (employeeName: string, db: Record<string, WeeklyEntry>) => {
  const entries = Object.values(db)
    .filter((entry) => entry.employeeName === employeeName)
    .sort((a, b) => parseIsoDateLocal(b.weekDate).getTime() - parseIsoDateLocal(a.weekDate).getTime());
  return entries[0];
};

const getWeekLabels = (startDateStr: string) => {
  const startDate = parseIsoDateLocal(startDateStr);
  return Array.from({ length: 4 }).map((_, i) => {
    const monday = new Date(startDate);
    monday.setDate(startDate.getDate() + i * 7);

    const friday = new Date(monday);
    friday.setDate(monday.getDate() + 4);

    const startDay = monday.getDate();
    const endDay = friday.getDate();
    const startMonth = monday.toLocaleDateString('en-GB', { month: 'long' });
    const endMonth = friday.toLocaleDateString('en-GB', { month: 'long' });
    const year = friday.getFullYear();

    if (startMonth === endMonth) {
      return `${startDay}–${endDay} ${startMonth} ${year}`;
    }
    return `${startDay} ${startMonth} – ${endDay} ${endMonth} ${year}`;
  });
};

const getAllLeaveDates = (startDateStr: string, annualLeave: boolean[][]) => {
  const startDate = parseIsoDateLocal(startDateStr);
  const leaveDates: Date[] = [];

  annualLeave.forEach((week, weekIdx) => {
    week.forEach((isOff, dayIdx) => {
      if (isOff) {
        const d = new Date(startDate);
        d.setDate(startDate.getDate() + weekIdx * 7 + dayIdx);
        leaveDates.push(d);
      }
    });
  });

  if (leaveDates.length === 0) return '-';

  const sorted = [...leaveDates].sort((a, b) => a.getTime() - b.getTime());
  const ranges: { start: Date; end: Date }[] = [];
  let currentRange: { start: Date; end: Date } | null = null;

  sorted.forEach((date) => {
    if (!currentRange) {
      currentRange = { start: date, end: date };
      return;
    }
    const nextDay = new Date(currentRange.end);
    nextDay.setDate(nextDay.getDate() + 1);
    if (date.toDateString() === nextDay.toDateString()) {
      currentRange.end = date;
    } else {
      ranges.push(currentRange);
      currentRange = { start: date, end: date };
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
      return;
    }

    if (startMonth !== currentMonthLabel) {
      if (currentMonthParts.length > 0) {
        resultParts.push(`${currentMonthParts.join(', ')} ${currentMonthLabel}`);
      }
      currentMonthParts = [];
      currentMonthLabel = startMonth;
    }

    if (startDay === endDay) currentMonthParts.push(`${startDay}`);
    else currentMonthParts.push(`${startDay}–${endDay}`);
  });

  if (currentMonthParts.length > 0) {
    resultParts.push(`${currentMonthParts.join(', ')} ${currentMonthLabel}`);
  }

  return resultParts.join(', ');
};

const MATTER_TYPE_ALIASES: Record<string, MatterType> = {
  Category1: 'Category1',
  Category2: 'Category2',
  Project: 'Project',
  'Category 1': 'Category1',
  'Category 2': 'Category2',
  'Category A': 'Category1',
  'Category B': 'Category2',
  'Category C': 'Project'
};

const coerceMatterType = (value: unknown): MatterType | null => {
  if (typeof value !== 'string') return null;
  return MATTER_TYPE_ALIASES[value.trim()] || null;
};

const getProjectType = (
  project: Pick<Project, 'category'> & Partial<Pick<Project, 'matterType'>>
): MatterType => coerceMatterType(project.category) ?? coerceMatterType(project.matterType) ?? 'Project';

const normalizeProject = (project: Project): Project => {
  const normalizedType = getProjectType(project);
  const rawCapacities = Array.isArray(project.capacities) ? project.capacities : [];
  const normalizedCapacities = [0, 1, 2, 3].map((index) =>
    normalizeCapacityInput(rawCapacities[index])
  ) as [number, number, number, number];
  return {
    ...project,
    category: normalizedType,
    matterType: normalizedType,
    tasks: typeof project.tasks === 'string' ? project.tasks : '',
    capacities: normalizedCapacities
  };
};

const normalizeWeeklyEntry = (entry: WeeklyEntry): WeeklyEntry => ({
  ...entry,
  interests: typeof entry.interests === 'string' ? entry.interests : '',
  projects: (entry.projects || []).map((project) => normalizeProject(project))
});

const normalizeDb = (source: Record<string, WeeklyEntry>): Record<string, WeeklyEntry> =>
  Object.fromEntries(
    Object.entries(source).map(([key, entry]) => [key, normalizeWeeklyEntry(entry)])
  ) as Record<string, WeeklyEntry>;

const getMatterTotals = (projects: Project[]) => {
  const totals = { Category1: 0, Category2: 0, Project: 0 } as Record<MatterType, number>;
  for (const p of projects) totals[getProjectType(p)] += 1;
  return totals;
};

const inferLoadBadgeClass = (load: number) => {
  if (load >= 100) return 'bg-red-100 text-red-700';
  if (load >= 80) return 'bg-orange-100 text-orange-700';
  if (load >= 40) return 'bg-blue-100 text-blue-700';
  return 'bg-green-100 text-green-700';
};

// --- UI Components ---

type ButtonProps = {
  children: React.ReactNode;
  onClick?: React.MouseEventHandler<HTMLButtonElement>;
  variant?: 'primary' | 'secondary' | 'danger' | 'ghost';
  className?: string;
  icon?: React.ComponentType<{ size?: number; className?: string }>;
  type?: 'button' | 'submit' | 'reset';
  disabled?: boolean;
};

const Button = ({
  children,
  onClick,
  variant = 'primary',
  className = '',
  icon: Icon,
  type = 'button',
  disabled = false
}: ButtonProps) => {
  const baseStyle =
    'flex items-center justify-center gap-2 px-4 py-2 rounded-md font-medium transition-all duration-200 text-sm disabled:opacity-50 disabled:cursor-not-allowed focus:outline-none focus:ring-2 focus:ring-[#133958] focus:ring-offset-1';
  const variants = {
    primary: 'bg-[#133958] text-white hover:bg-[#003866] shadow-sm',
    secondary: 'bg-white text-slate-700 border border-[#94adae] hover:bg-[#EAF0F0] shadow-sm',
    danger: 'bg-red-50 text-red-600 hover:bg-red-100 border border-red-200',
    ghost: 'text-slate-700 hover:bg-[#EAF0F0]'
  };

  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      className={`${baseStyle} ${variants[variant as keyof typeof variants]} ${className}`}
    >
      {Icon && <Icon size={16} />}
      {children}
    </button>
  );
};

const Card = ({
  children,
  className = '',
  style
}: {
  children: React.ReactNode;
  className?: string;
  style?: React.CSSProperties;
}) => (
  <div className={`bg-white rounded-xl shadow-sm border border-[#94adae] ${className}`} style={style}>
    {children}
  </div>
);

const MultiSelect = ({
  options,
  value = [],
  onChange,
  placeholder,
  disabled,
  invalid = false
}: any) => {
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
        className={`w-full p-2 border rounded-md text-sm bg-white focus-within:ring-2 flex justify-between items-center min-h-[38px] ${
          invalid
            ? 'border-red-400 focus-within:ring-red-500 focus-within:border-red-500'
            : 'border-slate-300 focus-within:ring-blue-500 focus-within:border-blue-500'
        } ${
          disabled ? 'bg-[#f5f8f8] cursor-not-allowed opacity-70' : 'cursor-pointer'
        }`}
        onClick={() => !disabled && setIsOpen(!isOpen)}
      >
        <div className="flex flex-wrap gap-1">
          {value.length === 0 ? (
            <span className="text-slate-500">{placeholder}</span>
          ) : (
            value.map((v: string) => (
              <span
                key={v}
                className="bg-[#f5f8f8] text-slate-700 text-xs px-2 py-0.5 rounded-full border border-slate-200"
              >
                {v}
              </span>
            ))
          )}
        </div>
        <ChevronDown
          size={16}
          className={`text-slate-400 transition-transform ${isOpen ? 'rotate-180' : ''}`}
        />
      </div>

      {isOpen && !disabled && (
        <div className="absolute z-50 w-full mt-1 bg-white border border-slate-200 rounded-md shadow-lg max-h-60 overflow-auto animate-in fade-in zoom-in-95 duration-100">
          {options.map((opt: string) => {
            const isSelected = value.includes(opt);
            return (
              <div
                key={opt}
                className={`flex items-center px-3 py-2 cursor-pointer text-sm ${
                  isSelected ? 'bg-blue-50 text-blue-700' : 'text-slate-700 hover:bg-[#f5f8f8]'
                }`}
                onClick={() => toggleOption(opt)}
              >
                <div
                  className={`w-4 h-4 mr-2 rounded border flex items-center justify-center transition-colors ${
                    isSelected ? 'bg-blue-600 border-blue-600' : 'border-slate-300 bg-white'
                  }`}
                >
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

function ListManager({
  title,
  items,
  newItem,
  setNewItem,
  onAdd,
  onRemove
}: {
  title: string;
  items: string[];
  newItem: string;
  setNewItem: React.Dispatch<React.SetStateAction<string>>;
  onAdd: () => void;
  onRemove: (item: string) => void;
}) {
  const shouldScroll = items.length > 4;

  return (
    <Card className="p-6 h-full flex flex-col">
      <h3 className="font-bold text-[#133958] mb-4">{title}</h3>
      <div className="flex gap-2 mb-4">
        <input
          className="flex-1 p-2 border border-slate-300 rounded-md text-sm bg-white"
          placeholder={`Add ${title}...`}
          value={newItem}
          onChange={(e) => setNewItem(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && onAdd()}
        />
        <Button onClick={onAdd} variant="secondary" className="px-3">
          <Plus size={16} />
        </Button>
      </div>
      <div className="flex-1 overflow-hidden rounded-lg border border-[#94adae]">
        <div className={shouldScroll ? 'max-h-60 overflow-y-auto' : 'overflow-y-hidden'}>
          <table className="w-full text-left text-sm text-slate-600">
            <thead className="sticky top-0 z-[1] bg-[#f5f8f8] border-b border-[#94adae]">
              <tr className="text-center">
                <th className="px-4 py-3 font-semibold text-[#133958]">
                  {title.slice(0, -1) || 'Name'}
                </th>
                <th className="px-4 py-3 font-semibold text-right text-[#133958]">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {items.length === 0 ? (
                <tr>
                  <td className="px-4 py-3 text-center text-slate-400" colSpan={2}>
                    No entries
                  </td>
                </tr>
              ) : (
                items.map((item: string) => (
                  <tr key={item}>
                    <td className="px-4 py-3 font-medium text-slate-900">{item}</td>
                    <td className="px-4 py-3 text-right">
                      <button
                        onClick={() => onRemove(item)}
                        className="text-red-500 hover:text-red-700 font-medium text-xs"
                      >
                        Remove
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </Card>
  );
}

// --- Main Application ---

export default function CapacityTrackerApp() {
  const [view, setView] = useState<'login' | 'employee' | 'management' | 'operations' | 'settings'>('login');
  const [currentUser, setCurrentUser] = useState<string>('');
  const [showSaveModal, setShowSaveModal] = useState(false);

  const [db, setDb] = useState<Record<string, WeeklyEntry>>(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY_DB);
      if (raw) return normalizeDb(JSON.parse(raw) as Record<string, WeeklyEntry>);
    } catch {
      // ignore
    }
    return normalizeDb(MOCK_DB);
  });

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_DB, JSON.stringify(db));
    } catch {
      // ignore
    }
  }, [db]);

  // App Settings State
  const [offices, setOffices] = useState<string[]>(INITIAL_OFFICES);
  const [mentors, setMentors] = useState<string[]>(INITIAL_MENTORS);
  const [languages, setLanguages] = useState<string[]>(INITIAL_LANGUAGES);
  const [employees, setEmployees] = useState<Employee[]>(INITIAL_EMPLOYEES);

  const currentEmployeeEntry = getLatestEntryForEmployee(currentUser, db);

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

  const upsertEntry = (entry: WeeklyEntry) => {
    const updatedEntry = {
      ...normalizeWeeklyEntry(entry),
      lastUpdated: new Date().toISOString()
    };
    const key = `${updatedEntry.employeeName}-${updatedEntry.weekDate}`;
    setDb((prev) => ({ ...prev, [key]: updatedEntry }));
    return updatedEntry;
  };

  // Explicit save (button)
  const handleSave = (entry: WeeklyEntry) => {
    upsertEntry(entry);
    setShowSaveModal(true);
  };

  // Auto-save (silent)
  const handleAutoSave = (entry: WeeklyEntry) => {
    upsertEntry(entry);
  };

  return (
    <div className="min-h-screen bg-[#f5f8f8] text-slate-900 font-sans selection:bg-[#94adae] selection:text-white">
      {/* Header */}
      <header className="sticky top-0 z-10 flex flex-col shadow-sm">
        <div className="bg-[#003866] py-2 flex justify-center items-center text-xs text-white/90">
          <span className="uppercase tracking-wider font-medium">Company Internal</span>
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
                <span className="text-sm text-slate-500 bg-[#f5f8f8] px-3 py-1 rounded-full border border-slate-200 hidden sm:inline-block">
                  Logged in as <span className="font-semibold text-[#133958]">{currentUser}</span>
                </span>
                <button
                  onClick={handleLogout}
                  className="flex items-center gap-2 px-4 py-2 rounded-md font-medium transition-all duration-200 text-sm text-[#133958] hover:bg-[#f5f8f8] border border-transparent hover:border-[#94adae]"
                >
                  <LogOut size={16} />
                  <span className="hidden sm:inline">Logout</span>
                </button>
              </div>
            )}
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 py-8">
        {view === 'login' && (
          <LoginScreen onLogin={handleLogin} employees={employees} />
        )}

        {view === 'employee' && (
          <EmployeeForm
            user={currentUser}
            initialData={currentEmployeeEntry}
            onSave={handleSave}
            onAutoSave={handleAutoSave}
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
          <TeamDashboard
            db={db}
            offices={offices}
            languages={languages}
            mentors={mentors}
          />
        )}

        {view === 'settings' && (
          <SettingsView
            offices={offices}
            setOffices={setOffices}
            mentors={mentors}
            setMentors={setMentors}
            languages={languages}
            setLanguages={setLanguages}
            employees={employees}
            setEmployees={setEmployees}
            onBack={handleLogout}
          />
        )}
      </main>

      {showSaveModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#133958]">Save Entry</h3>
            <p className="mt-2 text-sm text-slate-600">Entry saved successfully.</p>
            <div className="mt-5 flex justify-end gap-2">
              <Button variant="secondary" onClick={() => setShowSaveModal(false)}>
                Close
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// --- Settings View ---

function SettingsView({
  offices,
  setOffices,
  mentors,
  setMentors,
  languages,
  setLanguages,
  employees,
  setEmployees,
  onBack
}: any) {
  type PendingSettingsDelete = {
    title: string;
    message: string;
    onConfirm: () => void;
  };

  const [newOffice, setNewOffice] = useState('');
  const [newMentor, setNewMentor] = useState('');
  const [newLanguage, setNewLanguage] = useState('');
  const [newEmpName, setNewEmpName] = useState('');
  const [newEmpPass, setNewEmpPass] = useState('');
  const [employeeError, setEmployeeError] = useState('');
  const [pendingDelete, setPendingDelete] = useState<PendingSettingsDelete | null>(null);
  const [passwordModalEmployee, setPasswordModalEmployee] = useState<Employee | null>(null);
  const [itPasswordInput, setItPasswordInput] = useState('');
  const [isPasswordModalAuthorized, setIsPasswordModalAuthorized] = useState(false);
  const [passwordDraft, setPasswordDraft] = useState('');
  const [passwordDraftVisible, setPasswordDraftVisible] = useState(false);
  const [passwordModalError, setPasswordModalError] = useState('');

  const addItem = (list: string[], setList: any, item: string, setItem: any) => {
    if (item.trim() && !list.includes(item.trim())) {
      setList([...list, item.trim()]);
      setItem('');
    }
  };

  const requestDelete = (title: string, message: string, onConfirm: () => void) => {
    setPendingDelete({ title, message, onConfirm });
  };

  const cancelDelete = () => {
    setPendingDelete(null);
  };

  const confirmDelete = () => {
    if (!pendingDelete) return;
    pendingDelete.onConfirm();
    setPendingDelete(null);
  };

  const removeItem = (setList: any, item: string, label: string) => {
    requestDelete('Delete Item', `Are you sure you wish to delete "${item}" from ${label}?`, () => {
      setList((prev: string[]) => prev.filter((i: string) => i !== item));
    });
  };

  const addEmployee = () => {
    const trimmedName = newEmpName.trim();
    const trimmedPassword = newEmpPass.trim();

    if (!trimmedName || !trimmedPassword) {
      setEmployeeError('Name and password are required.');
      return;
    }

    const hasDuplicate = employees.some(
      (employee: Employee) => employee.name.toLowerCase() === trimmedName.toLowerCase()
    );
    if (hasDuplicate) {
      setEmployeeError('An employee with that name already exists.');
      return;
    }

    setEmployees([...employees, { name: trimmedName, password: trimmedPassword }]);
    setNewEmpName('');
    setNewEmpPass('');
    setEmployeeError('');
  };

  const removeEmployee = (name: string) => {
    requestDelete('Delete Employee', `Are you sure you wish to delete "${name}"?`, () => {
      setEmployees((prev: Employee[]) => prev.filter((e: Employee) => e.name !== name));
      setEmployeeError('');
    });
  };

  const openPasswordModal = (employee: Employee) => {
    setPasswordModalEmployee(employee);
    setItPasswordInput('');
    setIsPasswordModalAuthorized(false);
    setPasswordDraft(employee.password);
    setPasswordDraftVisible(false);
    setPasswordModalError('');
  };

  const closePasswordModal = () => {
    setPasswordModalEmployee(null);
    setItPasswordInput('');
    setIsPasswordModalAuthorized(false);
    setPasswordDraft('');
    setPasswordDraftVisible(false);
    setPasswordModalError('');
  };

  const verifyPasswordModalAccess = () => {
    if (itPasswordInput !== IT_MASTER_PASSWORD) {
      setPasswordModalError('Invalid IT password.');
      return;
    }
    setIsPasswordModalAuthorized(true);
    setPasswordModalError('');
  };

  const savePasswordChanges = () => {
    if (!passwordModalEmployee) return;
    const nextPassword = passwordDraft.trim();
    if (!nextPassword) {
      setPasswordModalError('Password cannot be empty.');
      return;
    }

    setEmployees((prev: Employee[]) =>
      prev.map((employee) =>
        employee.name === passwordModalEmployee.name
          ? { ...employee, password: nextPassword }
          : employee
      )
    );
    closePasswordModal();
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center gap-4">
        <button
          onClick={onBack}
          className="flex items-center text-sm text-[#678b8c] hover:text-[#133958] font-medium transition-colors"
        >
          <ArrowLeft size={16} className="mr-1" /> Back
        </button>
        <h2 className="text-2xl font-bold text-[#133958]">IT Configuration</h2>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {/* Employee Management */}
        <Card className="p-6 md:col-span-2">
          <h3 className="font-bold text-[#133958] mb-4 flex items-center gap-2">
            <Users size={20} /> Employee Management
          </h3>

          <div className="flex flex-col md:flex-row gap-2 mb-4 p-4 bg-[#f5f8f8] rounded-lg border border-[#94adae]">
            <input
              className="flex-1 p-2 border border-slate-300 rounded-md text-sm bg-white"
              placeholder="Employee Name"
              value={newEmpName}
              onChange={(e) => setNewEmpName(e.target.value)}
            />
            <input
              type="password"
              className="flex-1 p-2 border border-slate-300 rounded-md text-sm bg-white"
              placeholder="Set Password"
              value={newEmpPass}
              onChange={(e) => setNewEmpPass(e.target.value)}
            />
            <Button onClick={addEmployee}>Add Employee</Button>
          </div>

          {employeeError && <p className="mb-4 text-xs text-red-600">{employeeError}</p>}

          <div className="overflow-x-auto">
            <div className="overflow-hidden rounded-lg border border-[#94adae]">
              <table className="w-full text-left text-sm text-slate-600">
                <thead className="bg-[#f5f8f8] border-b border-[#94adae]">
                  <tr className="text-center">
                    <th className="px-4 py-3 font-semibold text-[#133958]">Name</th>
                    <th className="px-4 py-3 font-semibold text-[#133958]">Password</th>
                    <th className="px-4 py-3 font-semibold text-right text-[#133958]">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {employees.map((emp: Employee) => (
                    <tr key={emp.name}>
                      <td className="px-4 py-3 font-medium text-slate-900">{emp.name}</td>
                      <td className="px-4 py-3">
                        <div className="flex items-center justify-center gap-2">
                          <span className="font-mono text-slate-700">*****</span>
                          <button
                            type="button"
                            onClick={() => openPasswordModal(emp)}
                            className="text-[#678b8c] hover:text-[#133958] transition-colors"
                            title="View / edit password"
                          >
                            <Eye size={14} />
                          </button>
                        </div>
                      </td>
                      <td className="px-4 py-3 text-right">
                        <button
                          onClick={() => removeEmployee(emp.name)}
                          className="text-red-500 hover:text-red-700 font-medium text-xs"
                        >
                          Remove
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </Card>

        {/* Lists Management */}
        <ListManager
          title="Offices"
          items={offices}
          newItem={newOffice}
          setNewItem={setNewOffice}
          onAdd={() => addItem(offices, setOffices, newOffice, setNewOffice)}
          onRemove={(item: string) => removeItem(setOffices, item, 'offices')}
        />
        <ListManager
          title="Mentors"
          items={mentors}
          newItem={newMentor}
          setNewItem={setNewMentor}
          onAdd={() => addItem(mentors, setMentors, newMentor, setNewMentor)}
          onRemove={(item: string) => removeItem(setMentors, item, 'mentors')}
        />
        <ListManager
          title="Languages"
          items={languages}
          newItem={newLanguage}
          setNewItem={setNewLanguage}
          onAdd={() => addItem(languages, setLanguages, newLanguage, setNewLanguage)}
          onRemove={(item: string) => removeItem(setLanguages, item, 'languages')}
        />
      </div>

      {pendingDelete && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#133958]">{pendingDelete.title}</h3>
            <p className="mt-2 text-sm text-slate-600">{pendingDelete.message}</p>
            <div className="mt-5 flex justify-end gap-2">
              <Button variant="secondary" onClick={cancelDelete}>
                Cancel
              </Button>
              <Button variant="danger" onClick={confirmDelete}>
                Delete
              </Button>
            </div>
          </div>
        </div>
      )}

      {passwordModalEmployee && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#133958]">View / Edit Password</h3>
            <p className="mt-2 text-sm text-slate-600">
              Employee:{' '}
              <span className="font-semibold text-slate-800">{passwordModalEmployee.name}</span>
            </p>

            <div className="mt-4 space-y-3">
              <div>
                <label className="block text-xs font-semibold text-[#678b8c] uppercase tracking-wider">
                  IT Password
                </label>
                <input
                  type="password"
                  className="mt-1 w-full p-2 border border-slate-300 rounded-md text-sm bg-white"
                  value={itPasswordInput}
                  onChange={(e) => setItPasswordInput(e.target.value)}
                  disabled={isPasswordModalAuthorized}
                  placeholder="Enter IT password"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-[#678b8c] uppercase tracking-wider">
                  Password
                </label>
                <div className="mt-1 flex items-center gap-2">
                  <input
                    type={passwordDraftVisible ? 'text' : 'password'}
                    className={`w-full p-2 border border-slate-300 rounded-md text-sm bg-white ${
                      !isPasswordModalAuthorized ? 'bg-slate-50 text-slate-500 cursor-not-allowed' : ''
                    }`}
                    value={passwordDraft}
                    onChange={(e) => setPasswordDraft(e.target.value)}
                    disabled={!isPasswordModalAuthorized}
                    placeholder="Employee password"
                  />
                  <button
                    type="button"
                    onClick={() => setPasswordDraftVisible((prev) => !prev)}
                    className={`text-[#678b8c] hover:text-[#133958] transition-colors ${
                      !isPasswordModalAuthorized ? 'opacity-50 cursor-not-allowed' : ''
                    }`}
                    disabled={!isPasswordModalAuthorized}
                    title={passwordDraftVisible ? 'Hide password' : 'Show password'}
                  >
                    {passwordDraftVisible ? <EyeOff size={16} /> : <Eye size={16} />}
                  </button>
                </div>
              </div>

              {passwordModalError && (
                <p className="text-xs text-red-600">{passwordModalError}</p>
              )}
            </div>

            <div className="mt-5 flex justify-end gap-2">
              <Button variant="secondary" onClick={closePasswordModal}>
                Cancel
              </Button>
              {!isPasswordModalAuthorized ? (
                <Button variant="primary" onClick={verifyPasswordModalAccess}>
                  Verify
                </Button>
              ) : (
                <Button variant="primary" onClick={savePasswordChanges}>
                  Save
                </Button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// --- Sub-Views ---

function LoginScreen({
  onLogin,
  employees
}: {
  onLogin: (action: LoginAction) => void;
  employees: Employee[];
}) {
  const [step, setStep] = useState<'select' | 'auth'>('select');
  const [role, setRole] = useState<'employee' | 'management' | 'operations' | 'it' | null>(null);
  const [selectedEmployee, setSelectedEmployee] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

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
      if (password === IT_MASTER_PASSWORD) onLogin({ type: 'it' });
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
      const emp = employees.find((e) => e.name === selectedEmployee);
      if (emp && password === emp.password)
        onLogin({ type: 'employee', employeeName: selectedEmployee });
      else setError('Invalid Password');
    }
  };

  if (step === 'select') {
    return (
      <div className="max-w-md mx-auto mt-20 animate-in fade-in slide-in-from-bottom-4 duration-500">
        <Card className="p-8 space-y-6 text-center relative">
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
                  <ChevronRight
                    size={16}
                    className="text-[#678b8c] opacity-0 group-hover:opacity-100 transition-opacity"
                  />
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
                  <ChevronRight
                    size={16}
                    className="text-[#678b8c] opacity-0 group-hover:opacity-100 transition-opacity"
                  />
                </div>
                <p className="text-xs text-[#678b8c]">
                  View weekly busy and available Employees
                </p>
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
                  <ChevronRight
                    size={16}
                    className="text-[#678b8c] opacity-0 group-hover:opacity-100 transition-opacity"
                  />
                </div>
                <p className="text-xs text-[#678b8c]">View full management overview</p>
              </div>
            </button>
          </div>

          <p className="text-xs text-[#94adae] mt-6">
            Please contact IT if you have any questions
          </p>
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
          <div className="mx-auto w-12 h-12 rounded-full flex items-center justify-center mb-3 bg-[#EAF0F0]">
            {role === 'it' ? (
              <Settings size={24} className="text-[#133958]" />
            ) : (
              <Lock size={24} className="text-[#133958]" />
            )}
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
              <label className="text-xs font-semibold text-[#678b8c] uppercase tracking-wider">
                Employee Name
              </label>
              <select
                className="w-full p-2.5 bg-white border border-[#94adae] rounded-lg text-sm focus:ring-2 focus:ring-[#133958] outline-none transition-all"
                value={selectedEmployee}
                onChange={(e) => setSelectedEmployee(e.target.value)}
              >
                <option value="" disabled>
                  Select an Employee...
                </option>
                {employees.map((emp) => (
                  <option key={emp.name} value={emp.name}>
                    {emp.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div className="space-y-1.5">
            <label className="text-xs font-semibold text-[#678b8c] uppercase tracking-wider">
              Password
            </label>
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

          <Button type="submit" className="w-full justify-center mt-2">
            Login
          </Button>

          <div className="text-center mt-4">
            <p className="text-xs text-[#94adae]">Use your assigned credentials.</p>
          </div>
        </form>
      </Card>
    </div>
  );
}

function EmployeeForm({
  user,
  initialData,
  onSave,
  onAutoSave,
  readOnly = false,
  onBack,
  offices,
  mentors,
  languages
}: any) {
  const currentWeekStart = useMemo(() => getCurrentWeekStart(), []);
  const defaultData = useMemo<WeeklyEntry>(
    () => ({
      weekDate: currentWeekStart,
      employeeName: user,
      office: offices[0],
      mentor: '',
      languages: ['English'],
      interests: '',
      annualLeave: Array(4)
        .fill(null)
        .map(() => Array(5).fill(false)),
      availability2Weeks: 'With Capacity',
      capacityComments: Array(4).fill(''),
      lastUpdated: new Date().toISOString(),
      projects: []
    }),
    [offices, user, currentWeekStart]
  );

  const [formData, setFormData] = useState<WeeklyEntry>(
    normalizeWeeklyEntry(initialData || defaultData)
  );
  const [projectPendingDelete, setProjectPendingDelete] = useState<{ id: string; name: string } | null>(null);
  const [profileValidationTouched, setProfileValidationTouched] = useState(false);
  const [capacityInputDrafts, setCapacityInputDrafts] = useState<Record<string, string>>({});
  const sidebarColumnRef = useRef<HTMLDivElement | null>(null);
  const [mattersCardHeight, setMattersCardHeight] = useState<number | null>(null);
  const [requiredFieldIssues, setRequiredFieldIssues] = useState<string[] | null>(null);

  useEffect(() => {
    setFormData(normalizeWeeklyEntry(initialData || defaultData));
  }, [initialData, defaultData]);

  useEffect(() => {
    setCapacityInputDrafts({});
  }, [initialData, defaultData]);

  useEffect(() => {
    if (readOnly) return;
    setFormData((prev) =>
      prev.weekDate === currentWeekStart ? prev : { ...prev, weekDate: currentWeekStart }
    );
  }, [currentWeekStart, readOnly]);

  useEffect(() => {
    const sidebarEl = sidebarColumnRef.current;
    if (!sidebarEl) return;

    let rafId: number | null = null;
    const syncMattersHeight = () => {
      if (window.innerWidth < 1024) {
        setMattersCardHeight(null);
        return;
      }
      const nextHeight = Math.ceil(sidebarEl.getBoundingClientRect().height);
      setMattersCardHeight((prev) => (prev === nextHeight ? prev : nextHeight));
    };

    const scheduleSync = () => {
      if (rafId !== null) window.cancelAnimationFrame(rafId);
      rafId = window.requestAnimationFrame(syncMattersHeight);
    };

    scheduleSync();
    const resizeObserver = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(scheduleSync) : null;
    resizeObserver?.observe(sidebarEl);
    window.addEventListener('resize', scheduleSync);

    return () => {
      if (rafId !== null) window.cancelAnimationFrame(rafId);
      resizeObserver?.disconnect();
      window.removeEventListener('resize', scheduleSync);
    };
  }, []);

  const hasOfficeSelected = !!formData.office && formData.office.trim() !== '';
  const hasMentorSelected =
    !!formData.mentor && formData.mentor.trim() !== '' && formData.mentor !== 'Select Mentor';
  const hasLanguagesSelected = Array.isArray(formData.languages) && formData.languages.length > 0;
  const isMatterNameValid = (project: Project) => !!project.name && project.name.trim() !== '';
  const isMatterCategoryValid = (project: Project) => !!project.category && project.category.trim() !== '';
  const isMatterSupervisorValid = (project: Project) => !!project.owner && project.owner.trim() !== '';
  const areMattersValid = formData.projects.every(
    (project) =>
      isMatterNameValid(project) &&
      isMatterCategoryValid(project) &&
      isMatterSupervisorValid(project)
  );
  const isProfileValid = hasOfficeSelected && hasMentorSelected && hasLanguagesSelected;
  const isFormValidForSave = isProfileValid && areMattersValid;

  // Silent autosave (debounced)
  useEffect(() => {
    if (readOnly || !isFormValidForSave) return;
    const t = window.setTimeout(() => {
      onAutoSave?.(formData);
    }, 900);
    return () => window.clearTimeout(t);
  }, [formData, readOnly, onAutoSave, isFormValidForSave]);

  const weekCommencingDate = readOnly ? formData.weekDate : currentWeekStart;
  const weekLabels = getWeekLabels(weekCommencingDate);
  const capacityComments = formData.capacityComments || Array(4).fill('');

  // Weekly totals: project capacities + leave
  const weeklyTotals = [0, 1, 2, 3].map((i) => {
    const projectSumPct = formData.projects.reduce((sum, p) => sum + (p.capacities[i] || 0), 0);
    const leaveCount = (formData.annualLeave[i] || []).filter(Boolean).length;
    const projectHours = (projectSumPct / 100) * HOURS_PER_WEEK;
    const leaveHours = leaveCount * LEAVE_HOURS_PER_DAY;
    return Math.max(0, Math.round((projectHours + leaveHours) * 60) / 60);
  });

  const getStatusColor = (hours: number) => {
    if (hours > HOURS_PER_WEEK) return 'text-red-600 bg-red-50 border-red-200';
    if (hours >= 32) return 'text-orange-600 bg-orange-50 border-orange-200';
    if (hours >= 16) return 'text-blue-600 bg-blue-50 border-blue-200';
    return 'text-green-600 bg-green-50 border-green-200';
  };

  const availabilitySelectClasses: Record<Availability2Weeks, string> = {
    'With Capacity': 'bg-emerald-50 text-emerald-700 border-emerald-300',
    'Limited Capacity': 'bg-amber-50 text-amber-700 border-amber-300',
    'No Capacity': 'bg-red-50 text-red-700 border-red-300',
    'Over Capacity': 'bg-red-100 text-red-800 border-red-400'
  };

  const sortProjectsForSave = (projects: Project[]) => {
    const categoryOrder: Record<Project['category'], number> = {
      'Category1': 0,
      'Category2': 1,
      'Project': 2
    };

    return [...projects].sort((a, b) => {
      const categoryDiff = categoryOrder[getProjectType(a)] - categoryOrder[getProjectType(b)];
      if (categoryDiff !== 0) return categoryDiff;

      for (let weekIdx = 0; weekIdx < 4; weekIdx += 1) {
        const loadDiff = (b.capacities[weekIdx] || 0) - (a.capacities[weekIdx] || 0);
        if (loadDiff !== 0) return loadDiff;
      }

      const totalLoadDiff =
        b.capacities.reduce((sum, load) => sum + load, 0) -
        a.capacities.reduce((sum, load) => sum + load, 0);
      if (totalLoadDiff !== 0) return totalLoadDiff;

      return (a.name || '').localeCompare(b.name || '');
    });
  };

  const addProject = () => {
    const generatedId =
      typeof crypto !== 'undefined' && typeof (crypto as any).randomUUID === 'function'
        ? (crypto as any).randomUUID()
        : Math.random().toString(36).slice(2, 11);

    const newProject: Project = {
      id: generatedId,
      name: '',
      category: 'Project',
      matterType: 'Project',
      owner: '',
      tasks: '',
      capacities: [0, 0, 0, 0]
    };

    setFormData((prev) => ({
      ...prev,
      projects: [newProject, ...prev.projects]
    }));
  };

  const requestRemoveProject = (id: string, name: string) => {
    const projectLabel = name.trim() || 'Untitled Project';
    setProjectPendingDelete({ id, name: projectLabel });
  };

  const cancelRemoveProject = () => {
    setProjectPendingDelete(null);
  };

  const confirmRemoveProject = () => {
    if (!projectPendingDelete) return;
    setFormData((prev) => ({
      ...prev,
      projects: prev.projects.filter((p) => p.id !== projectPendingDelete.id)
    }));
    setProjectPendingDelete(null);
  };

  const dismissRequiredFieldsModal = () => {
    setRequiredFieldIssues(null);
  };

  const getRequiredFieldIssues = () => {
    const issues: string[] = [];

    if (!hasOfficeSelected) issues.push('Office is required.');
    if (!hasMentorSelected) issues.push('Mentor is required.');
    if (!hasLanguagesSelected) issues.push('Working Language(s) is required.');

    formData.projects.forEach((project, index) => {
      const matterLabel = `Matter ${index + 1}`;
      if (!isMatterNameValid(project)) issues.push(`${matterLabel}: Matter Name is required.`);
      if (!isMatterCategoryValid(project)) issues.push(`${matterLabel}: Category is required.`);
      if (!isMatterSupervisorValid(project)) issues.push(`${matterLabel}: Supervisor is required.`);
    });

    return issues;
  };

  const handleSaveEntry = () => {
    const issues = getRequiredFieldIssues();
    if (issues.length > 0) {
      setProfileValidationTouched(true);
      setRequiredFieldIssues(issues);
      return;
    }

    const sortedProjects = sortProjectsForSave(formData.projects);
    const nextData = { ...formData, projects: sortedProjects };
    setProfileValidationTouched(false);
    setRequiredFieldIssues(null);
    setFormData(nextData);
    onSave(nextData);
  };

  const updateProject = (id: string, field: keyof Project, value: any) => {
    setFormData({
      ...formData,
      projects: formData.projects.map((p) => {
        if (p.id !== id) return p;
        if (field === 'category') {
          const nextType = getProjectType({ category: value, matterType: p.matterType });
          return { ...p, category: nextType, matterType: nextType };
        }
        if (field === 'tasks' && typeof value === 'string') {
          return { ...p, tasks: limitToWordCount(value, COMMENT_WORD_LIMIT) };
        }
        return { ...p, [field]: value };
      })
    });
  };

  const updateCapacity = (id: string, index: number, value: number) => {
    setFormData({
      ...formData,
      projects: formData.projects.map((p) => {
        if (p.id !== id) return p;
        const newCapacities = [...p.capacities] as [number, number, number, number];
        newCapacities[index] = clampCapacity(value);
        return { ...p, capacities: newCapacities };
      })
    });
  };

  const getCapacityInputKey = (projectId: string, weekIndex: number) => `${projectId}-${weekIndex}`;

  const getCapacityInputValue = (projectId: string, weekIndex: number, capacityPercent: number) => {
    const inputKey = getCapacityInputKey(projectId, weekIndex);
    return capacityInputDrafts[inputKey] ?? formatHoursInput(percentToHours(capacityPercent));
  };

  const handleCapacityInputChange = (projectId: string, weekIndex: number, rawValue: string) => {
    const inputKey = getCapacityInputKey(projectId, weekIndex);
    setCapacityInputDrafts((prev) => ({ ...prev, [inputKey]: rawValue }));
    updateCapacity(projectId, weekIndex, hoursToPercent(normalizeHoursInput(rawValue)));
  };

  const handleCapacityInputBlur = (projectId: string, weekIndex: number) => {
    const inputKey = getCapacityInputKey(projectId, weekIndex);
    setCapacityInputDrafts((prev) => {
      if (!(inputKey in prev)) return prev;
      const { [inputKey]: _removed, ...rest } = prev;
      return rest;
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

  const updateCapacityComment = (weekIndex: number, value: string) => {
    if (readOnly) return;
    const nextComments = [...(formData.capacityComments || Array(4).fill(''))];
    nextComments[weekIndex] = limitToWordCount(value, COMMENT_WORD_LIMIT);
    setFormData({ ...formData, capacityComments: nextComments });
  };

  const moveProject = (id: string, direction: 'up' | 'down') => {
    if (readOnly) return;
    setFormData((prev) => {
      const idx = prev.projects.findIndex((p) => p.id === id);
      if (idx < 0) return prev;
      const targetIdx = direction === 'up' ? idx - 1 : idx + 1;
      if (targetIdx < 0 || targetIdx >= prev.projects.length) return prev;
      const next = [...prev.projects];
      const [item] = next.splice(idx, 1);
      next.splice(targetIdx, 0, item);
      return { ...prev, projects: next };
    });
  };

  const totals = getMatterTotals(formData.projects);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      {/* Header and Controls */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-end lg:items-stretch">
        <div className="lg:col-span-1 lg:flex lg:flex-col lg:justify-center">
          {readOnly && onBack && (
            <button
              onClick={onBack}
              className="flex items-center text-xs text-[#678b8c] hover:text-[#133958] mb-4 group transition-colors"
            >
              <ArrowLeft
                size={14}
                className="mr-1 group-hover:-translate-x-1 transition-transform"
              />
              Back to Dashboard
            </button>
          )}

          <h2 className="text-2xl font-bold text-[#133958]">
            {readOnly ? 'Employee Profile' : 'Monthly Capacity Form'}
          </h2>

          <div className="mt-2 space-y-1">
            <p className="text-[#678b8c]">
              Week commencing{' '}
              <span className="font-semibold text-slate-700">
                {formatDisplayDate(weekCommencingDate)}
              </span>
            </p>
            <p className="text-[11px] leading-snug text-[#94adae]">
              This information if for planning purposes only and to signal your expected capacity over
              the coming weeks.
            </p>
          </div>
        </div>

        <div className="lg:col-span-2 flex flex-col gap-3">
          <div className="grid grid-cols-1 md:grid-cols-3 items-center gap-2">
            <div className="hidden md:block" />
            <div className="text-center">
              <div className="text-sm font-semibold text-[#133958]">Total Workload</div>
              <div className="text-xs text-[#678b8c] whitespace-nowrap">
                Category1 Matters: <span className="font-semibold text-slate-700">{totals.Category1}</span>{' '}
                · Category2 Matters: <span className="font-semibold text-slate-700">{totals.Category2}</span> ·
                Projects: <span className="font-semibold text-slate-700">{totals.Project}</span>
              </div>
            </div>
            <div className="flex justify-center md:justify-end">
              {!readOnly && (
                <Button onClick={handleSaveEntry} icon={Save} className="whitespace-nowrap">
                  Save Entry
                </Button>
              )}
            </div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 w-full">
            {weeklyTotals.map((total, idx) => {
              const label = weekLabels[idx].replace(/ \d{4}$/, '');
              const colorClass = getStatusColor(total);
              const totalLabel = `${formatHoursInput(total)} ${total > 1 ? 'hrs' : 'hr'}`;

              return (
                <div
                  key={idx}
                  className={`flex flex-col items-center justify-center px-4 py-2 rounded-lg border ${colorClass}`}
                >
                  <span className="text-[10px] font-bold uppercase tracking-tight mb-0.5 opacity-80 text-center">
                    {label}
                  </span>
                  <div className="flex items-center gap-1.5">
                    <span
                      className="text-2xl font-bold leading-none cursor-help"
                      title="40 hours = 100%"
                      aria-label="40 hours = 100%"
                    >
                      {total > HOURS_PER_WEEK ? 'Over 40' : totalLabel}
                    </span>
                    {total > HOURS_PER_WEEK && <AlertCircle size={16} />}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 lg:items-stretch">
        {/* Sidebar */}
        <div ref={sidebarColumnRef} className="space-y-6 h-fit lg:col-span-1">
          {/* Profile Card */}
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold text-[#133958] flex items-center gap-2">
              <FileText size={18} className="text-[#678b8c]" /> Profile
            </h3>

            <div className="space-y-3">
              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">
                  Employee Name
                </label>
                <input
                  type="text"
                  className="w-full p-2 border border-slate-300 rounded-md text-sm bg-[#f5f8f8] text-slate-500 cursor-not-allowed"
                  value={formData.employeeName}
                  readOnly
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">
                  Office <span className="text-red-500">*</span>
                </label>
                <select
                  className={`w-full p-2 border rounded-md text-sm bg-white focus:ring-2 outline-none ${
                    profileValidationTouched && !hasOfficeSelected
                      ? 'border-red-400 focus:ring-red-500 focus:border-red-500'
                      : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                  } ${
                    readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''
                  }`}
                  value={formData.office}
                  onChange={(e) => setFormData({ ...formData, office: e.target.value })}
                  disabled={readOnly}
                >
                  {offices.map((opt: string) => (
                    <option key={opt}>{opt}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">
                  Mentor <span className="text-red-500">*</span>
                </label>
                <select
                  className={`w-full p-2 border rounded-md text-sm bg-white focus:ring-2 outline-none ${
                    profileValidationTouched && !hasMentorSelected
                      ? 'border-red-400 focus:ring-red-500 focus:border-red-500'
                      : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                  } ${
                    readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''
                  }`}
                  value={formData.mentor}
                  onChange={(e) => setFormData({ ...formData, mentor: e.target.value })}
                  disabled={readOnly}
                >
                  <option value="">Select Mentor</option>
                  {mentors.map((opt: string) => (
                    <option key={opt}>{opt}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">
                  Working Language(s) <span className="text-red-500">*</span>
                </label>
                <MultiSelect
                  options={languages}
                  value={formData.languages}
                  onChange={(val: string[]) => setFormData({ ...formData, languages: val })}
                  placeholder="Select Languages"
                  disabled={readOnly}
                  invalid={profileValidationTouched && !hasLanguagesSelected}
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-[#678b8c] mb-1">
                  Interests
                </label>
                <textarea
                  className={`w-full h-24 min-h-24 max-h-24 resize-none overflow-y-auto p-2 border border-slate-300 rounded-md text-xs bg-white ${
                    readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''
                  }`}
                  value={formData.interests || ''}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      interests: limitToWordCount(e.target.value, COMMENT_WORD_LIMIT)
                    })
                  }
                  disabled={readOnly}
                  placeholder="Share your interests, preferred work areas, or goals..."
                />
              </div>
            </div>
          </Card>

          {/* Capacity Summary Card */}
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold text-[#133958] flex items-center gap-2">
              <CalendarDays size={18} className="text-[#678b8c]" /> Capacity Summary
            </h3>

            <div>
              <label className="block text-xs font-medium text-[#678b8c] mb-1">
                Capacity in the next two weeks
              </label>
              <select
                className={`w-full p-2 border rounded-md text-sm font-medium ${
                  availabilitySelectClasses[formData.availability2Weeks]
                } ${readOnly ? 'opacity-70 cursor-not-allowed' : ''}`}
                value={formData.availability2Weeks}
                onChange={(e) =>
                  setFormData({ ...formData, availability2Weeks: e.target.value as Availability2Weeks })
                }
                disabled={readOnly}
              >
                <option className="text-slate-900 bg-white">With Capacity</option>
                <option className="text-slate-900 bg-white">Limited Capacity</option>
                <option className="text-slate-900 bg-white">No Capacity</option>
                <option className="text-slate-900 bg-white">Over Capacity</option>
              </select>
            </div>

            <div className="space-y-4">
              {weekLabels.map((dateStr, weekIdx) => {
                const extraHighlight = weekIdx < 2 ? 'ring-1 ring-blue-200 bg-blue-50/30' : '';

                return (
                  <div
                    key={weekIdx}
                    className={`bg-[#f5f8f8] p-3 rounded-md border border-slate-200 ${extraHighlight}`}
                  >
                    <div className="text-xs font-bold text-slate-700 mb-2 text-center">
                      Week of {dateStr.replace(/ \d{4}$/, '')}
                    </div>
                    <div className="mt-1">
                      <label className="block text-[10px] font-semibold text-[#678b8c] mb-1 uppercase tracking-wide">
                        Comments
                      </label>
                      <textarea
                        className={`w-full h-24 min-h-24 max-h-24 resize-none overflow-y-auto p-2 border border-slate-300 rounded-md text-xs bg-white ${
                          readOnly ? 'bg-[#f5f8f8] text-slate-500 cursor-not-allowed' : ''
                        }`}
                        value={capacityComments[weekIdx] || ''}
                        onChange={(e) => updateCapacityComment(weekIdx, e.target.value)}
                        disabled={readOnly}
                        placeholder="Any relevant notes regarding your capacity or workload this week..."
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>

          {/* Annual Leave Card */}
          <Card className="p-6 space-y-4">
            <h3 className="font-semibold text-[#133958] flex items-center gap-2">
              <CalendarDays size={18} className="text-[#678b8c]" /> Annual Leave
            </h3>

            <div className="space-y-4">
              {weekLabels.map((dateStr, weekIdx) => {
                const days = formData.annualLeave[weekIdx] || [false, false, false, false, false];
                const extraHighlight = weekIdx < 2 ? 'ring-1 ring-blue-200 bg-blue-50/30' : '';

                return (
                  <div
                    key={`annual-leave-week-${weekIdx}`}
                    className={`bg-[#f5f8f8] p-3 rounded-md border border-slate-200 ${extraHighlight}`}
                  >
                    <div className="text-xs font-bold text-slate-700 mb-2 text-center">
                      Week of {dateStr.replace(/ \d{4}$/, '')}
                    </div>
                    <div className="flex justify-between">
                      {WEEKDAYS.map((dayName, dayIdx) => (
                        <label
                          key={dayIdx}
                          className={`flex flex-col items-center gap-1 group ${
                            readOnly ? 'cursor-not-allowed opacity-70' : 'cursor-pointer'
                          }`}
                        >
                          <span className="text-[10px] text-slate-500 font-medium group-hover:text-blue-600">
                            {dayName}
                          </span>
                          <span
                            aria-hidden="true"
                            className={`w-4 h-4 rounded border flex items-center justify-center text-[11px] font-bold transition-colors ${
                              days[dayIdx]
                                ? 'bg-red-600 border-red-600 text-white'
                                : 'bg-white border-slate-300 text-transparent'
                            } ${readOnly ? 'cursor-not-allowed' : 'cursor-pointer'}`}
                          >
                            ×
                          </span>
                          <input
                            type="checkbox"
                            className="sr-only"
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

        {/* Matters List */}
        <Card
          className="p-6 lg:col-span-2 flex flex-col h-full min-h-0 overflow-hidden"
          style={mattersCardHeight ? { height: mattersCardHeight } : undefined}
        >
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3">
              <h3 className="font-semibold text-[#133958]">Matters</h3>
              <span className="bg-[#f5f8f8] text-slate-600 px-2.5 py-0.5 rounded-full text-xs font-semibold border border-slate-200">
                {formData.projects.length}
              </span>
            </div>
            <div className="flex gap-2">
              {!readOnly && (
                <Button variant="secondary" onClick={addProject} icon={Plus}>
                  Add Matter
                </Button>
              )}
            </div>
          </div>

          <div className="flex-1 min-h-0 overflow-y-auto pr-1 space-y-4">
            {formData.projects.length === 0 && (
              <div className="text-center py-10 text-slate-400 border-2 border-dashed border-slate-200 rounded-lg bg-[#f5f8f8]">
                No matters added yet.
              </div>
            )}

            {formData.projects.map((project, index) => (
              <div
                key={project.id}
                className="p-4 bg-slate-50 rounded-lg border border-[#94adae] group hover:border-[#678b8c] transition-colors space-y-4"
              >
                {/* Row 1: Basic Info */}
                <div className="grid grid-cols-12 gap-4 items-end">
                  <div className="col-span-12 md:col-span-4">
                    <label className="text-xs font-medium text-[#678b8c] ml-1">
                      Matter Name <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. project"
                      className={`w-full p-2 border rounded-md text-sm bg-white ${
                        profileValidationTouched && !isMatterNameValid(project)
                          ? 'border-red-400 focus:ring-2 focus:ring-red-500 focus:border-red-500'
                          : 'border-slate-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500'
                      } ${
                        readOnly ? 'bg-slate-50 text-slate-600 cursor-not-allowed' : ''
                      }`}
                      value={project.name}
                      onChange={(e) => updateProject(project.id, 'name', e.target.value)}
                      readOnly={readOnly}
                    />
                  </div>

                  <div className="col-span-12 md:col-span-3">
                    <label className="text-xs font-medium text-[#678b8c] ml-1">
                      Category <span className="text-red-500">*</span>
                    </label>
                    <select
                      className={`w-full p-2 border rounded-md text-sm bg-white focus:ring-2 outline-none ${
                        profileValidationTouched && !isMatterCategoryValid(project)
                          ? 'border-red-400 focus:ring-red-500 focus:border-red-500'
                          : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                      } ${
                        readOnly ? 'bg-slate-50 text-slate-600 cursor-not-allowed' : ''
                      }`}
                      value={project.category}
                      onChange={(e) => updateProject(project.id, 'category', e.target.value)}
                      disabled={readOnly}
                    >
                      {PROJECT_CATEGORIES.map((cat) => (
                        <option key={cat} value={cat}>
                          {cat}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="col-span-12 md:col-span-3">
                    <label className="text-xs font-medium text-[#678b8c] ml-1">
                      Supervisor <span className="text-red-500">*</span>
                    </label>
                    <select
                      className={`w-full p-2 border rounded-md text-sm bg-white focus:ring-2 outline-none ${
                        profileValidationTouched && !isMatterSupervisorValid(project)
                          ? 'border-red-400 focus:ring-red-500 focus:border-red-500'
                          : 'border-slate-300 focus:ring-blue-500 focus:border-blue-500'
                      } ${
                        readOnly ? 'bg-slate-50 text-slate-600 cursor-not-allowed' : ''
                      }`}
                      value={project.owner}
                      onChange={(e) => updateProject(project.id, 'owner', e.target.value)}
                      disabled={readOnly}
                    >
                      <option value="">Select Supervisor</option>
                      {project.owner && !mentors.includes(project.owner) && (
                        <option value={project.owner}>{project.owner}</option>
                      )}
                      {mentors.map((opt: string) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="col-span-12 md:col-span-2 flex items-center justify-end gap-2 pb-1">
                    {!readOnly && (
                      <>
                        <button
                          onClick={() => moveProject(project.id, 'up')}
                          className={`text-slate-400 hover:text-slate-700 transition-colors ${
                            index === 0 ? 'opacity-30 cursor-not-allowed' : ''
                          }`}
                          title="Move up"
                          disabled={index === 0}
                        >
                          <ArrowUp size={18} />
                        </button>
                        <button
                          onClick={() => moveProject(project.id, 'down')}
                          className={`text-slate-400 hover:text-slate-700 transition-colors ${
                            index === formData.projects.length - 1
                              ? 'opacity-30 cursor-not-allowed'
                              : ''
                          }`}
                          title="Move down"
                          disabled={index === formData.projects.length - 1}
                        >
                          <ArrowDown size={18} />
                        </button>
                        <button
                          onClick={() => requestRemoveProject(project.id, project.name)}
                          className="text-slate-400 hover:text-red-500 transition-colors ml-2"
                          title="Remove"
                        >
                          <Trash2 size={18} />
                        </button>
                      </>
                    )}
                  </div>
                </div>

                {/* Row 2: Capacity Forecast */}
                <div className="rounded-md border border-slate-300 pt-3 pb-3 px-3 mt-2 bg-[#f8fafc]">
                  <span className="block mb-2 text-xs font-bold text-[#678b8c]">
                    Workload Forecast (Hours)
                  </span>
                  <div className="grid grid-cols-4 gap-3">
                    {weekLabels.map((dateStr, idx) => (
                      <div key={idx}>
                        <label
                          className="text-[10px] uppercase font-bold mb-1.5 block truncate text-center text-[#133958]"
                          title={dateStr}
                        >
                          {dateStr.replace(/ \d{4}$/, '')}
                        </label>
                        <div className="relative">
                          <input
                            type="text"
                            inputMode="text"
                            pattern="[0-9:]*"
                            className={`w-full py-1.5 pl-10 pr-10 border rounded-md text-sm text-center border-slate-200 text-slate-600 bg-white focus:border-blue-400 transition-colors ${
                              readOnly ? 'cursor-not-allowed' : ''
                            }`}
                            value={getCapacityInputValue(project.id, idx, project.capacities[idx] || 0)}
                            onChange={(e) => handleCapacityInputChange(project.id, idx, e.target.value)}
                            onBlur={() => handleCapacityInputBlur(project.id, idx)}
                            readOnly={readOnly}
                            placeholder="0:00"
                            title="Enter hours as H:MM (e.g., 5:45)"
                          />
                          <span
                            aria-hidden="true"
                            className="pointer-events-none absolute inset-y-0 right-3 flex items-center text-xs font-medium text-slate-400"
                          >
                            hrs
                          </span>
                        </div>
                      </div>
                    ))}
                  </div>
                  <div className="mt-3">
                    <label className="block text-[10px] font-semibold text-[#678b8c] mb-1 uppercase tracking-wide">
                      Tasks
                    </label>
                    <textarea
                      className={`w-full h-24 min-h-24 max-h-24 resize-none overflow-y-auto p-2 border border-slate-300 rounded-md text-xs bg-white ${
                        readOnly ? 'bg-slate-50 text-slate-600 cursor-not-allowed' : ''
                      }`}
                      value={project.tasks || ''}
                      onChange={(e) => updateProject(project.id, 'tasks', e.target.value)}
                      readOnly={readOnly}
                      placeholder="Add task notes for this matter..."
                    />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>

      {requiredFieldIssues && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="w-full max-w-lg rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#133958]">Required Fields Missing</h3>
            <p className="mt-2 text-sm text-slate-600">
              Complete the required fields before saving.
            </p>
            <div className="mt-3 max-h-56 overflow-y-auto rounded-md border border-slate-200 bg-[#f8fafc] p-3">
              <ul className="list-disc pl-5 space-y-1 text-sm text-slate-700">
                {requiredFieldIssues.map((issue, idx) => (
                  <li key={`${issue}-${idx}`}>{issue}</li>
                ))}
              </ul>
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <Button variant="secondary" onClick={dismissRequiredFieldsModal}>
                OK
              </Button>
            </div>
          </div>
        </div>
      )}

      {projectPendingDelete && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#133958]">Delete Matter</h3>
            <p className="mt-2 text-sm text-slate-600">
              Are you sure you wish to delete "{projectPendingDelete.name}"?
            </p>
            <div className="mt-5 flex justify-end gap-2">
              <Button variant="secondary" onClick={cancelRemoveProject}>
                Cancel
              </Button>
              <Button variant="danger" onClick={confirmRemoveProject}>
                Delete
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function TeamDashboard({
  db,
  offices,
  languages,
  mentors
}: {
  db: Record<string, WeeklyEntry>;
  offices: string[];
  languages: string[];
  mentors: string[];
}) {
  return (
    <Dashboard
      db={db}
      offices={offices}
      languages={languages}
      mentors={mentors}
      showInsights={false}
      weekDrivenTable={true}
    />
  );
}

function TeamEmployeeProfile({ entry, onBack }: { entry: WeeklyEntry; onBack: () => void }) {
  const weekLabels = getWeekLabels(entry.weekDate);
  const weeklyLoads = Array.from({ length: 4 }).map((_, weekIdx) => {
    const projectLoad = entry.projects.reduce(
      (acc, project) => acc + (project.capacities[weekIdx] || 0),
      0
    );
    const leaveCount = (entry.annualLeave[weekIdx] || []).filter(Boolean).length;
    const leaveLoad = (leaveCount / WEEKDAYS.length) * 100;
    return clampWeekLoadPercent(projectLoad + leaveLoad);
  });
  const comments = entry.capacityComments || Array(4).fill('');
  const interestsText = entry.interests?.trim() ? entry.interests : '-';
  const annualLeaveByWeek = weekLabels.map((label, weekIdx) => {
    const weekLeave = entry.annualLeave[weekIdx] || [];
    return {
      weekIdx,
      weekLabel: label.replace(/ \d{4}$/, ''),
      leaveLabel: formatLeaveDaySpans(weekLeave)
    };
  });

  return (
    <div className="space-y-6">
      <Button variant="secondary" onClick={onBack} icon={ArrowLeft} className="w-fit">
        Back to Team Dashboard
      </Button>

      <Card className="p-6">
        <h2 className="text-2xl font-bold text-[#133958]">Employee Profile</h2>
        <div className="mt-4 grid grid-cols-1 gap-4 md:grid-cols-2">
          <div>
            <p className="text-xs uppercase tracking-wide text-[#678b8c]">Employee</p>
            <p className="text-base font-semibold text-slate-800">{entry.employeeName}</p>
          </div>
          <div>
            <p className="text-xs uppercase tracking-wide text-[#678b8c]">Office</p>
            <p className="text-base font-semibold text-slate-800">{entry.office || '-'}</p>
          </div>
          <div>
            <p className="text-xs uppercase tracking-wide text-[#678b8c]">Languages</p>
            <p className="text-base font-semibold text-slate-800">
              {entry.languages.join(' / ') || '-'}
            </p>
          </div>
          <div>
            <p className="text-xs uppercase tracking-wide text-[#678b8c]">Status</p>
            <p className="text-base font-semibold text-slate-800">{entry.availability2Weeks}</p>
          </div>
          <div className="md:col-span-2">
            <p className="text-xs uppercase tracking-wide text-[#678b8c]">Interests</p>
            <p className="text-base font-semibold text-slate-800 whitespace-pre-wrap break-words">
              {interestsText}
            </p>
          </div>
        </div>
      </Card>

      <Card className="p-6">
        <h3 className="text-lg font-semibold text-[#133958]">Weekly Capacity and Comments</h3>
        <div className="mt-4 overflow-hidden rounded-lg border border-[#94adae]">
          <table className="w-full table-fixed text-left text-sm text-slate-600">
            <colgroup>
              <col className="w-[28%]" />
              <col className="w-[16%]" />
              <col className="w-[56%]" />
            </colgroup>
            <thead className="bg-[#f5f8f8] border-b border-[#94adae] uppercase text-xs font-semibold text-[#133958]">
              <tr>
                <th className="px-4 py-3 text-center">Week</th>
                <th className="px-4 py-3 text-center">Capacity</th>
                <th className="px-4 py-3">Deadline(s) / Comments</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {weekLabels.map((label, weekIdx) => (
                <tr key={`profile-week-row-${weekIdx}`}>
                  <td className="px-4 py-3 text-center font-semibold text-slate-700">
                    Week {weekIdx + 1} ({label.replace(/ \d{4}$/, '')})
                  </td>
                  <td className="px-4 py-3 text-center">
                    <span
                      className={`inline-block px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                        weeklyLoads[weekIdx]
                      )}`}
                    >
                      {weeklyLoads[weekIdx]}%
                    </span>
                  </td>
                  <td className="px-4 py-3 text-slate-700">
                    {comments[weekIdx]?.trim() || '-'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card className="p-6">
        <h3 className="text-lg font-semibold text-[#133958]">Annual Leave by Week</h3>
        <div className="mt-4 grid grid-cols-1 gap-3 md:grid-cols-2 xl:grid-cols-4">
          {annualLeaveByWeek.map((week) => (
            <div
              key={`profile-annual-leave-${week.weekIdx}`}
              className="rounded-lg border border-[#94adae] bg-[#f5f8f8] p-3"
            >
              <p className="text-xs font-bold uppercase tracking-wide text-[#678b8c]">
                Week {week.weekIdx + 1} ({week.weekLabel})
              </p>
              <p className="mt-2 text-sm font-semibold text-slate-700 break-words">
                {week.leaveLabel}
              </p>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

function Dashboard({
  db,
  offices,
  languages,
  mentors,
  showInsights = true,
  weekDrivenTable = false
}: {
  db: Record<string, WeeklyEntry>;
  offices: string[];
  languages: string[];
  mentors: string[];
  showInsights?: boolean;
  weekDrivenTable?: boolean;
}) {
  type DashboardEntry = WeeklyEntry & {
    weeklyLoads: number[];
    averageLoad: number;
    loadDelta: number;
    weekLoad1: number;
    weekLoad2: number;
    weekLoad3: number;
    weekLoad4: number;
    totalCategory1: number;
    totalCategory2: number;
    totalProjects: number;
  };

  const [sortConfig, setSortConfig] = useState<{ key: string; direction: 'asc' | 'desc' } | null>(
    null
  );
  const [activeWeek, setActiveWeek] = useState<number | null>(weekDrivenTable ? null : 0);
  const [peopleLoadMode, setPeopleLoadMode] = useState<'most' | 'least'>('most');
  const [selectedEntry, setSelectedEntry] = useState<WeeklyEntry | null>(null);

  const normalizeEmployeeKey = (name: string) => name.trim().toLowerCase();

  const latestEntries = useMemo(() => {
    const byEmployee = new Map<string, WeeklyEntry>();

    Object.values(db).forEach((entry) => {
      const normalizedKey = normalizeEmployeeKey(entry.employeeName || '');
      const existing = byEmployee.get(normalizedKey);
      if (!existing) {
        byEmployee.set(normalizedKey, entry);
        return;
      }

      const existingWeek = parseIsoDateLocal(existing.weekDate).getTime();
      const nextWeek = parseIsoDateLocal(entry.weekDate).getTime();
      if (nextWeek > existingWeek) {
        byEmployee.set(normalizedKey, entry);
        return;
      }

      if (nextWeek === existingWeek && entry.lastUpdated > existing.lastUpdated) {
        byEmployee.set(normalizedKey, entry);
      }
    });

    return Array.from(byEmployee.values());
  }, [db]);

  const referenceWeekDate = getCurrentWeekStart();
  const weekLabels = getWeekLabels(referenceWeekDate);
  const relativeWeekLabels = ['This week', 'Next week', 'Week 3', 'Week 4'];
  const spanStartDate = parseIsoDateLocal(referenceWeekDate);
  const spanEndDate = new Date(spanStartDate);
  spanEndDate.setDate(spanStartDate.getDate() + 25);
  const dateSpanLabel = `${spanStartDate.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  })} to ${spanEndDate.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  })}`;
  const selectedWeekIndex = activeWeek ?? 0;
  const activeWeekKey = `weekLoad${selectedWeekIndex + 1}`;

  const getProjectBadgeClass = (load: number) => {
    if (load >= 50) return 'bg-red-100 text-red-700 border border-red-200';
    if (load >= 25) return 'bg-orange-100 text-orange-700 border border-orange-200';
    if (load > 0) return 'bg-blue-50 text-blue-700 border border-blue-200';
    return 'bg-[#f5f8f8] text-[#678b8c] border border-[#94adae]';
  };

  const tableData = useMemo<DashboardEntry[]>(() => {
    return latestEntries.map((entry) => {
      const weeklyLoads = Array.from({ length: 4 }).map((_, weekIdx) => { 
        const projectLoad = entry.projects.reduce(
          (acc, project) => acc + (project.capacities[weekIdx] || 0),
          0
        );
        const leaveCount = (entry.annualLeave[weekIdx] || []).filter(Boolean).length;
        const leaveLoad = (leaveCount / WEEKDAYS.length) * 100;
        return clampWeekLoadPercent(projectLoad + leaveLoad);
      });

      const averageLoad = Math.round(
        weeklyLoads.reduce((acc, load) => acc + load, 0) / weeklyLoads.length
      );
      const totals = getMatterTotals(entry.projects);

      return {
        ...entry,
        weeklyLoads,
        averageLoad,
        loadDelta: weeklyLoads[3] - weeklyLoads[0],
        weekLoad1: weeklyLoads[0],
        weekLoad2: weeklyLoads[1],
        weekLoad3: weeklyLoads[2],
        weekLoad4: weeklyLoads[3],
        totalCategory1: totals.Category1,
        totalCategory2: totals.Category2,
        totalProjects: totals.Project
      };
    });
  }, [latestEntries]);

  const filteredData = useMemo(() => tableData, [tableData]);

  const weeklySummary = useMemo(() => {
    return Array.from({ length: 4 }).map((_, weekIdx) => {
      const weekLoads = filteredData.map((entry) => entry.weeklyLoads[weekIdx] || 0);
      const totalLoad = weekLoads.reduce((acc, load) => acc + load, 0);
      const avgLoad = weekLoads.length > 0 ? Math.round(totalLoad / weekLoads.length) : 0;
      const withCapacity = weekLoads.filter((load) => load < 80).length;
      const atOrOverCapacity = weekLoads.filter((load) => load >= 80).length;
      const totalLeaveDays = filteredData.reduce(
        (acc, entry) => acc + (entry.annualLeave[weekIdx] || []).filter(Boolean).length,
        0
      );
      const avgLeaveDays = weekLoads.length > 0 ? totalLeaveDays / weekLoads.length : 0;

      return { avgLoad, withCapacity, atOrOverCapacity, avgLeaveDays: avgLeaveDays.toFixed(1) };
    });
  }, [filteredData]);

  const activeWeekInsights = useMemo(() => {
    const uniqueEntriesByEmployee = new Map<string, (typeof filteredData)[number]>();
    filteredData.forEach((entry) => {
      const normalizedKey = normalizeEmployeeKey(entry.employeeName || '');
      const existing = uniqueEntriesByEmployee.get(normalizedKey);
      if (!existing) {
        uniqueEntriesByEmployee.set(normalizedKey, entry);
        return;
      }

      const existingWeek = parseIsoDateLocal(existing.weekDate).getTime();
      const nextWeek = parseIsoDateLocal(entry.weekDate).getTime();
      if (nextWeek > existingWeek || (nextWeek === existingWeek && entry.lastUpdated > existing.lastUpdated)) {
        uniqueEntriesByEmployee.set(normalizedKey, entry);
      }
    });

    const uniqueEntries = Array.from(uniqueEntriesByEmployee.values());
    const weekLoads = uniqueEntries.map((entry) => entry.weeklyLoads[selectedWeekIndex] || 0);
    const averageLoad =
      weekLoads.length > 0
        ? Math.round(weekLoads.reduce((acc, load) => acc + load, 0) / weekLoads.length)
        : 0;
    const atCapacityCount = weekLoads.filter((load) => load >= 80 && load < 100).length;
    const overCapCount = weekLoads.filter((load) => load >= 100).length;
    const lookingForWorkCount = weekLoads.filter((load) => load < 80).length;

    const mostLoadedPeople = [...uniqueEntries]
      .sort(
        (a, b) =>
          (b.weeklyLoads[selectedWeekIndex] || 0) - (a.weeklyLoads[selectedWeekIndex] || 0)
      )
      .slice(0, 3)
      .map((entry) => ({
        name: entry.employeeName.trim(),
        load: entry.weeklyLoads[selectedWeekIndex] || 0
      }));

    const leastLoadedPeople = [...uniqueEntries]
      .sort(
        (a, b) =>
          (a.weeklyLoads[selectedWeekIndex] || 0) - (b.weeklyLoads[selectedWeekIndex] || 0)
      )
      .slice(0, 3)
      .map((entry) => ({
        name: entry.employeeName.trim(),
        load: entry.weeklyLoads[selectedWeekIndex] || 0
      }));

    const projectTotals = new Map<string, number>();
    uniqueEntries.forEach((entry) => {
      entry.projects.forEach((project) => {
        const load = project.capacities[selectedWeekIndex] || 0;
        if (load <= 0) return;
        const key = project.name || '(Untitled)';
        projectTotals.set(key, (projectTotals.get(key) || 0) + load);
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
  }, [filteredData, selectedWeekIndex]);

  const sortedData = useMemo(() => {
    const appliedSort = sortConfig || { key: activeWeekKey, direction: 'desc' as const };
    return [...filteredData].sort((a, b) => {
      const aValue: any = (a as any)[appliedSort.key];
      const bValue: any = (b as any)[appliedSort.key];

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
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') direction = 'desc';
    setSortConfig({ key, direction });
  };

  const getSortIcon = (name: string) => {
    if (!sortConfig || sortConfig.key !== name) return <div className="w-4" />;
    return sortConfig.direction === 'asc' ? (
      <ArrowUp size={14} className="text-slate-500" />
    ) : (
      <ArrowDown size={14} className="text-slate-500" />
    );
  };

  const getWeekLeaveDaysLabel = (weekLeave: boolean[]) => {
    return formatLeaveDaySpans(weekLeave);
  };

  if (selectedEntry) {
    if (weekDrivenTable) {
      return <TeamEmployeeProfile entry={selectedEntry} onBack={() => setSelectedEntry(null)} />;
    }
    return (
      <EmployeeForm
        user={selectedEntry.employeeName}
        initialData={selectedEntry}
        onSave={() => {}}
        onAutoSave={() => {}}
        readOnly={true}
        onBack={() => setSelectedEntry(null)}
        offices={offices}
        languages={languages}
        mentors={mentors}
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
      </div>

      <div className="grid grid-cols-1 gap-3 md:grid-cols-2 xl:grid-cols-4">
        {weeklySummary.map((summary, weekIdx) => {
          const isActive = activeWeek === weekIdx;
          const shortLabel = weekLabels[weekIdx]?.replace(/ \d{4}$/, '');

          return (
            <button
              key={`week-summary-${weekIdx}`}
              type="button"
              onClick={() =>
                setActiveWeek((prev) =>
                  weekDrivenTable ? (prev === weekIdx ? null : weekIdx) : weekIdx
                )
              }
              className={`text-left rounded-xl border p-4 transition-colors ${
                isActive
                  ? 'border-[#133958] bg-[#e6eff3]'
                  : 'border-[#94adae] bg-white hover:bg-[#f5f8f8]'
              }`}
            >
              <div className="flex items-center justify-between">
                <p className="text-sm font-semibold text-[#133958]">
                  {relativeWeekLabels[weekIdx] || `Week ${weekIdx + 1}`}
                </p>
                <span
                  className={`px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                    summary.avgLoad
                  )}`}
                >
                  {summary.avgLoad}%
                </span>
              </div>
              <p className="mt-1 text-xs text-[#678b8c]">{shortLabel}</p>
              <p className="mt-3 text-xs text-slate-600">
                {summary.withCapacity} with capacity | {summary.atOrOverCapacity} at or over
                {' '}capacity
              </p>
              {!weekDrivenTable && (
                <p className="mt-1 text-xs text-slate-500">
                  Avg leave: {summary.avgLeaveDays} days
                </p>
              )}
            </button>
          );
        })}
      </div>

      {showInsights && (
        <div className="grid grid-cols-1 gap-4 xl:grid-cols-3">
          <Card className="p-4">
            <p className="text-sm font-semibold text-[#133958]">
              Active snapshot (Week {selectedWeekIndex + 1})
            </p>
            <p className="mt-2 text-3xl font-bold text-[#133958]">
              {activeWeekInsights.averageLoad}%
            </p>
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
                {peopleLoadMode === 'most' ? 'Busiest Workload' : 'Most Available'} (Week{' '}
                {selectedWeekIndex + 1})
              </p>
              <div className="inline-flex rounded-md border border-slate-200 bg-white p-0.5">
                <button
                  type="button"
                  onClick={() => setPeopleLoadMode('most')}
                  className={`px-2 py-1 text-xs rounded ${
                    peopleLoadMode === 'most'
                      ? 'bg-[#133958] text-white'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  Most
                </button>
                <button
                  type="button"
                  onClick={() => setPeopleLoadMode('least')}
                  className={`px-2 py-1 text-xs rounded ${
                    peopleLoadMode === 'least'
                      ? 'bg-[#133958] text-white'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  Least
                </button>
              </div>
            </div>
            <div className="mt-3 space-y-2">
              {(peopleLoadMode === 'most'
                ? activeWeekInsights.mostLoadedPeople
                : activeWeekInsights.leastLoadedPeople
              ).map((person) => (
                <div key={`${person.name}-${person.load}`} className="flex items-center justify-between text-sm">
                  <span className="text-slate-700">{person.name}</span>
                  <span
                    className={`px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                      person.load
                    )}`}
                  >
                    {person.load}%
                  </span>
                </div>
              ))}
            </div>
          </Card>

          <Card className="p-4">
            <p className="text-sm font-semibold text-[#133958]">
              Top matters by demand (Week {selectedWeekIndex + 1})
            </p>
            <div className="mt-3 space-y-2">
              {activeWeekInsights.topProjects.length === 0 && (
                <p className="text-sm text-slate-500">No scheduled load in this week.</p>
              )}
              {activeWeekInsights.topProjects.map((project) => (
                <div key={project.name} className="flex items-center justify-between text-sm">
                  <span className="text-slate-700 truncate pr-4">{project.name}</span>
                  <span
                    className={`px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                      project.total
                    )}`}
                  >
                    {project.total}%
                  </span>
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {weekDrivenTable ? (
        <Card className="overflow-hidden">
          <div className="pb-2 overflow-x-auto">
            <table
              className="w-full table-fixed text-sm text-slate-600 min-w-[1150px]"
            >
              <thead className="bg-[#f5f8f8] border-b border-[#94adae] uppercase text-xs font-semibold text-[#133958] text-left">
                <tr className="text-left h-14">
                  <th
                    className="px-4 py-2 cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('employeeName')}
                  >
                    <div className="flex items-center justify-start gap-1">
                      Employee {getSortIcon('employeeName')}
                    </div>
                  </th>
                  <th
                    className="px-4 py-2 cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('office')}
                  >
                    <div className="flex items-center justify-start gap-1">
                      Office {getSortIcon('office')}
                    </div>
                  </th>
                  {activeWeek !== null && <th className="px-4 py-2">Languages</th>}
                  <th
                    className="px-4 py-2 cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('availability2Weeks')}
                  >
                    <div className="flex items-center justify-start gap-1">
                      Availability Outlook {getSortIcon('availability2Weeks')}
                    </div>
                  </th>
                  {activeWeek !== null ? (
                    <th
                      className="px-4 py-2 cursor-pointer hover:bg-slate-100 select-none"
                      onClick={() => requestSort(`weekLoad${activeWeek + 1}`)}
                    >
                      <div className="flex items-center justify-start gap-1">
                        Week {activeWeek + 1} {getSortIcon(`weekLoad${activeWeek + 1}`)}
                      </div>
                    </th>
                  ) : (
                    weekLabels.map((_, weekIdx) => (
                      <th
                        key={`team-week-col-${weekIdx}`}
                        className="px-4 py-2 cursor-pointer hover:bg-slate-100 select-none"
                        onClick={() => requestSort(`weekLoad${weekIdx + 1}`)}
                      >
                        <div className="flex items-center justify-start gap-1">
                          Week {weekIdx + 1} {getSortIcon(`weekLoad${weekIdx + 1}`)}
                        </div>
                      </th>
                    ))
                  )}
                  <th
                    className="px-4 py-2 w-[78px] max-w-[78px] text-center cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('totalCategory1')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Category1 {getSortIcon('totalCategory1')}
                    </div>
                  </th>
                  <th
                    className="px-4 py-2 w-[78px] max-w-[78px] text-center cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('totalCategory2')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Category2 {getSortIcon('totalCategory2')}
                    </div>
                  </th>
                  <th
                    className="px-4 py-2 w-[78px] max-w-[78px] text-center cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('totalProjects')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Projects {getSortIcon('totalProjects')}
                    </div>
                  </th>
                  {activeWeek !== null && (
                    <th className="px-4 py-2 text-center whitespace-nowrap">
                      Annual Leave (Week {activeWeek + 1})
                    </th>
                  )}
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100">
                {sortedData.map((entry) => (
                  <tr
                    key={`${entry.employeeName}-${entry.weekDate}`}
                    className="hover:bg-slate-50 transition-colors cursor-pointer"
                    onClick={() => setSelectedEntry(entry)}
                  >
                    <td className="px-4 py-3 font-medium text-slate-900 text-left">
                      {entry.employeeName}
                    </td>
                    <td className="px-4 py-3 text-left">{entry.office || '-'}</td>
                    {activeWeek !== null && (
                      <td className="px-4 py-3 text-left text-xs">
                        {entry.languages.join(' / ') || '-'}
                      </td>
                    )}
                    <td className="px-4 py-3 text-left">{entry.availability2Weeks}</td>
                    {activeWeek !== null ? (
                      <td className="px-4 py-3 text-left">
                        <span
                          className={`px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                            entry.weeklyLoads[activeWeek]
                          )}`}
                        >
                          {entry.weeklyLoads[activeWeek]}%
                        </span>
                      </td>
                    ) : (
                      entry.weeklyLoads.map((load, weekIdx) => (
                        <td key={`team-load-${entry.employeeName}-${weekIdx}`} className="px-4 py-3 text-left">
                          <span
                            className={`px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                              load
                            )}`}
                          >
                            {load}%
                          </span>
                        </td>
                      ))
                    )}
                    <td className="px-4 py-3 text-center text-slate-600 font-medium w-[78px] max-w-[78px]">
                      <span className="block truncate" title={String(entry.totalCategory1)}>
                        {entry.totalCategory1}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-center text-slate-600 font-medium w-[78px] max-w-[78px]">
                      <span className="block truncate" title={String(entry.totalCategory2)}>
                        {entry.totalCategory2}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-center text-slate-600 font-medium w-[78px] max-w-[78px]">
                      <span className="block truncate" title={String(entry.totalProjects)}>
                        {entry.totalProjects}
                      </span>
                    </td>
                    {activeWeek !== null && (
                      <td className="px-4 py-3 text-center text-xs text-slate-600">
                        {getWeekLeaveDaysLabel(entry.annualLeave[activeWeek] || [])}
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      ) : (
        <Card className="overflow-hidden">
          <div className="pb-2 overflow-x-auto">
            <table className="w-full table-fixed text-sm text-slate-600 min-w-[1100px]">
              <thead className="bg-[#f5f8f8] border-b border-[#94adae] uppercase text-xs font-semibold text-[#133958] text-center">
                <tr>
                  <th
                    className="px-3 py-2 cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('employeeName')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Employee {getSortIcon('employeeName')}
                    </div>
                  </th>
                  <th
                    className="px-3 py-2 cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('office')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Office {getSortIcon('office')}
                    </div>
                  </th>
                  <th
                    className="px-3 py-2 cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('availability2Weeks')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Availability Outlook {getSortIcon('availability2Weeks')}
                    </div>
                  </th>
                  {weekLabels.map((_, weekIdx) => (
                    <th
                      key={`week-col-${weekIdx}`}
                      className="px-3 py-2 w-[96px] max-w-[96px] cursor-pointer hover:bg-slate-100 select-none"
                      onClick={() => requestSort(`weekLoad${weekIdx + 1}`)}
                    >
                      <div className="flex items-center justify-center gap-1">
                        Week {weekIdx + 1} {getSortIcon(`weekLoad${weekIdx + 1}`)}
                      </div>
                    </th>
                  ))}
                  <th
                    className="px-3 py-2 w-[78px] max-w-[78px] cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('totalCategory1')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Category1 {getSortIcon('totalCategory1')}
                    </div>
                  </th>
                  <th
                    className="px-3 py-2 w-[78px] max-w-[78px] cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('totalCategory2')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Category2 {getSortIcon('totalCategory2')}
                    </div>
                  </th>
                  <th
                    className="px-3 py-2 w-[78px] max-w-[78px] cursor-pointer hover:bg-slate-100 select-none"
                    onClick={() => requestSort('totalProjects')}
                  >
                    <div className="flex items-center justify-center gap-1">
                      Projects {getSortIcon('totalProjects')}
                    </div>
                  </th>
                  <th className="px-3 py-2 w-[14%]">Annual Leave</th>
                  <th className="px-3 py-2 w-[22%]">Top Matters (Week {selectedWeekIndex + 1})</th>
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100">
                {sortedData.map((entry) => {
                  const topProjects = [...entry.projects]
                    .filter((project) => (project.capacities[selectedWeekIndex] || 0) > 0)
                    .sort(
                      (a, b) =>
                        (b.capacities[selectedWeekIndex] || 0) -
                        (a.capacities[selectedWeekIndex] || 0)
                    )
                    .slice(0, 4);

                  return (
                    <tr
                      key={`${entry.employeeName}-${entry.weekDate}`}
                      className="hover:bg-slate-50 transition-colors cursor-pointer"
                      onClick={() => setSelectedEntry(entry)}
                    >
                      <td className="px-3 py-3 font-medium text-slate-900 text-center">
                        {entry.employeeName}
                      </td>
                      <td className="px-3 py-3 text-center">{entry.office || '-'}</td>
                      <td className="px-3 py-3 text-center">{entry.availability2Weeks}</td>
                      {entry.weeklyLoads.map((load, weekIdx) => (
                        <td
                          key={`load-${entry.employeeName}-${weekIdx}`}
                          className="px-3 py-3 w-[96px] max-w-[96px] text-center"
                        >
                          <span
                            className={`px-2 py-1 rounded-full text-xs font-bold ${inferLoadBadgeClass(
                              load
                            )}`}
                          >
                            {load}%
                          </span>
                        </td>
                      ))}
                      <td className="px-3 py-3 text-center text-slate-600 font-medium w-[78px] max-w-[78px]">
                        <span className="block truncate" title={String(entry.totalCategory1)}>
                          {entry.totalCategory1}
                        </span>
                      </td>
                      <td className="px-3 py-3 text-center text-slate-600 font-medium w-[78px] max-w-[78px]">
                        <span className="block truncate" title={String(entry.totalCategory2)}>
                          {entry.totalCategory2}
                        </span>
                      </td>
                      <td className="px-3 py-3 text-center text-slate-600 font-medium w-[78px] max-w-[78px]">
                        <span className="block truncate" title={String(entry.totalProjects)}>
                          {entry.totalProjects}
                        </span>
                      </td>
                      <td className="px-3 py-3 text-xs text-slate-600 text-center break-words">
                        {getAllLeaveDates(entry.weekDate, entry.annualLeave)}
                      </td>
                      <td className="px-3 py-3 align-top">
                        {topProjects.length === 0 ? (
                          <p className="text-xs text-slate-500 text-center">
                            No matters this week
                          </p>
                        ) : (
                          <div className="flex flex-col gap-2">
                            {topProjects.map((project) => {
                              const cap = project.capacities[selectedWeekIndex] || 0;
                              const displayName = project.name || '(Untitled)';

                              return (
                                <div
                                  key={project.id}
                                  className="flex flex-col text-xs w-full p-1.5 bg-[#f5f8f8] rounded border border-[#94adae]"
                                >
                                  <div className="flex justify-between items-center mb-0.5">
                                    <span
                                      className="font-medium text-[#133958] truncate w-[70%]"
                                      title={displayName}
                                    >
                                      {displayName}
                                    </span>
                                    <span
                                      className={`px-1.5 py-px rounded text-[10px] font-bold ${getProjectBadgeClass(
                                        cap
                                      )}`}
                                    >
                                      {cap.toFixed(0)}%
                                    </span>
                                  </div>
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
      )}
    </div>
  );
}
