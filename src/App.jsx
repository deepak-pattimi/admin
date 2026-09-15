import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { io } from 'socket.io-client';
import * as XLSX from 'xlsx';

const Loader = () => (
  <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', padding: '60px', animation: 'fadeIn 0.3s ease-out' }}>
    <svg width="44" height="44" viewBox="0 0 50 50" style={{ animation: 'spin 1s linear infinite' }}>
      <circle cx="25" cy="25" r="20" fill="none" stroke="rgba(148, 163, 184, 0.2)" strokeWidth="4"></circle>
      <circle cx="25" cy="25" r="20" fill="none" stroke="#38bdf8" strokeWidth="4" strokeDasharray="31.4 100" strokeLinecap="round"></circle>
    </svg>
    <div style={{ marginTop: '16px', color: '#94a3b8', fontSize: '14px', fontWeight: '500', letterSpacing: '0.5px' }}>LOADING DATA...</div>
    <style>{`@keyframes spin { 100% { transform: rotate(360deg); } } @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }`}</style>
  </div>
);

function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(() => localStorage.getItem('adminAuth') === 'true');
  const [loggedInEmail, setLoggedInEmail] = useState(() => localStorage.getItem('adminEmail') || '');
  const [showPassword, setShowPassword] = useState(false);
  const [authForm, setAuthForm] = useState({ email: '', password: '' });
  const [authError, setAuthError] = useState('');

  const [activeTab, setActiveTab] = useState(() => localStorage.getItem('adminActiveTab') || 'employees');
  const [leaves, setLeaves] = useState([]);
  const [employees, setEmployees] = useState([]);
  const [liveData, setLiveData] = useState([]);
  const [newEmployee, setNewEmployee] = useState({ name: '', email: '' });
  const [isAddingEmployee, setIsAddingEmployee] = useState(false);
  const [editingEmployeeId, setEditingEmployeeId] = useState(null);

  // Detailed View State
  const [selectedEmployeeId, setSelectedEmployeeId] = useState(null);
  const [detailedDate, setDetailedDate] = useState(new Date().toISOString().split('T')[0]);
  const [detailedData, setDetailedData] = useState(null);
  const [activeSubModal, setActiveSubModal] = useState(null); // 'appActivity' | 'appUsage' | null
  const [lastUpdate, setLastUpdate] = useState(Date.now());

  // Attendance Log State
  const [attendanceDate, setAttendanceDate] = useState(new Date().toISOString().split('T')[0]);
  const [dailyAttendanceData, setDailyAttendanceData] = useState([]);
  const [reportMonth, setReportMonth] = useState(new Date().toISOString().substring(0, 7));

  // Salary Tab State
  const [salaryMonth, setSalaryMonth] = useState(new Date().toISOString().substring(0, 7)); // e.g. "2026-08"
  const [salaryReport, setSalaryReport] = useState([]);
  const [loadingSalary, setLoadingSalary] = useState(false);
  const [salarySearch, setSalarySearch] = useState('');

  // Calendar Tab State
  const [calendarMonth, setCalendarMonth] = useState(new Date().toISOString().substring(0, 7));
  const [holidays, setHolidays] = useState([]);
  const [loadingCalendar, setLoadingCalendar] = useState(false);

  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  const [editFormData, setEditFormData] = useState({ newId: '', name: '', email: '', password: '' });
  const [processingLeaves, setProcessingLeaves] = useState({});
  const [uiTick, setUiTick] = useState(0);

  const [loadingLeaves, setLoadingLeaves] = useState(true);
  const [loadingLive, setLoadingLive] = useState(true);
  const [loadingAttendance, setLoadingAttendance] = useState(true);

  const pendingLeavesCount = leaves.filter(l => l.status === 'PENDING').length;

  useEffect(() => {
    localStorage.setItem('adminActiveTab', activeTab);
  }, [activeTab]);

  const fetchData = async () => {
    if (employees.length === 0) setLoadingLeaves(true);
    try {
      console.log(`[FRONTEND LOG] Fetching data from API URL: ${import.meta.env.VITE_API_URL}`);
      const leavesRes = await axios.get(`${import.meta.env.VITE_API_URL}/api/leaves`);
      setLeaves(Array.isArray(leavesRes.data) ? leavesRes.data : []);
      const empRes = await axios.get(`${import.meta.env.VITE_API_URL}/api/employees`);
      console.log('[FRONTEND LOG] Received Employees Data:', empRes.data);
      setEmployees(Array.isArray(empRes.data) ? empRes.data : []);
    } catch (err) {
      console.error('Error fetching data', err);
    } finally {
      setLoadingLeaves(false);
    }
  };

  const fetchLiveData = async () => {
    if (liveData.length === 0) setLoadingLive(true);
    try {
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/attendance/live`);
      setLiveData(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('Failed to fetch live data', err);
    } finally {
      setLoadingLive(false);
    }
  };

  const fetchDetailedData = async (empId, dateStr) => {
    try {
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/attendance/employee/${empId}?date=${dateStr}`);
      setDetailedData(res.data);
    } catch (err) {
      console.error('Failed to fetch detailed data', err);
    }
  };

  const autoFitColumns = (ws, data) => {
    if (!data || data.length === 0) return;
    const objectKeys = Object.keys(data[0] || {});
    ws['!cols'] = objectKeys.map(key => {
      let maxLen = key.length;
      data.forEach(row => {
        const valStr = row[key] !== undefined && row[key] !== null ? String(row[key]) : '';
        if (valStr.length > maxLen) {
          maxLen = valStr.length;
        }
      });
      return { wch: Math.max(maxLen + 4, 14) };
    });
  };

  const downloadTodaysReport = () => {
    if (dailyAttendanceData.length === 0) return alert('No data to download for today.');
    const exportData = dailyAttendanceData.map(emp => ({
      'Employee Name': emp.name,
      'Employee ID': emp.employeeId,
      'Department': emp.department,
      'Clock In': emp.clockIn ? new Date(emp.clockIn).toLocaleTimeString() : '--',
      'Clock Out': emp.clockOut ? new Date(emp.clockOut).toLocaleTimeString() : (emp.clockIn ? 'Active' : '--'),
      'Total Active (hrs)': Number((emp.totalMinutes / 60).toFixed(2)),
      'Status': emp.totalMinutes >= 420 ? 'Present' : (emp.totalMinutes >= 60 ? 'Requirement Not Met' : (emp.onLeave ? 'On Leave' : 'Absent'))
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    autoFitColumns(ws, exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Today Report");
    XLSX.writeFile(wb, `Daily_Attendance_${attendanceDate}.xlsx`);
  };

  const downloadMonthlyReport = async () => {
    try {
      const monthStr = attendanceDate.substring(0, 7); // e.g. "2026-07"
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/attendance/monthly?month=${monthStr}`);
      const data = res.data;
      if (data.length === 0) return alert('No data to download for this month.');
      
      const exportData = data.map(row => {
        const isLate = row.clockIn && new Date(row.clockIn).getHours() >= 10;
        let status = 'Present';
        if (row.totalMinutes < 420) status = 'Requirement Not Met';
        if (!row.clockIn || row.totalMinutes < 60) status = 'Absent';

        return {
          'Date': row.date,
          'Employee Name': row.name,
          'Employee ID': row.employeeId,
          'Department': row.department,
          'Clock In': row.clockIn ? new Date(row.clockIn).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '--',
          'Clock Out': row.clockOut ? new Date(row.clockOut).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '--',
          'Total Active (hrs)': Number((row.totalMinutes / 60).toFixed(2)),
          'Status': isLate ? `${status} (Late)` : status,
          'App Usage Breakdown': row.appUsageStr
        };
      });

      const ws = XLSX.utils.json_to_sheet(exportData);
      autoFitColumns(ws, exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Monthly Report");
      XLSX.writeFile(wb, `Monthly_Attendance_${monthStr}.xlsx`);
    } catch (err) {
      console.error('Failed to download monthly report', err);
      alert('Failed to download monthly report');
    }
  };

  const downloadSimpleMonthlyReport = async () => {
    try {
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/attendance/monthly?month=${reportMonth}`);
      const data = res.data;
      if (data.length === 0) return alert('No data to download for this month.');
      
      const exportData = data.map(row => {
        const isLate = row.clockIn && new Date(row.clockIn).getHours() >= 10;
        let status = 'Present';
        if (row.totalMinutes < 420) status = 'Requirement Not Met';
        if (!row.clockIn || row.totalMinutes < 60) status = 'Absent';

        return {
          'Date': row.date,
          'Employee Name': row.name,
          'Employee ID': row.employeeId,
          'Department': row.department,
          'Clock In': row.clockIn ? new Date(row.clockIn).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '--',
          'Clock Out': row.clockOut ? new Date(row.clockOut).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '--',
          'Total Active (hrs)': Number((row.totalMinutes / 60).toFixed(2)),
          'Status': isLate ? `${status} (Late)` : status
        };
      });

      const ws = XLSX.utils.json_to_sheet(exportData);
      autoFitColumns(ws, exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Simple Monthly Report");
      XLSX.writeFile(wb, `Simple_Monthly_Attendance_${reportMonth}.xlsx`);
    } catch (err) {
      console.error('Failed to download simple monthly report', err);
      alert('Failed to download simple monthly report');
    }
  };

  const downloadSalaryReport = () => {
    if (!salaryReport || salaryReport.length === 0) {
      return alert('No salary data available to download for this month.');
    }
    
    const exportData = salaryReport.map(emp => ({
      'Employee Name': emp.name,
      'Total Days': emp.totalDays,
      'Days Present': emp.daysPresent,
      'Days Absent': emp.daysAbsent,
      'Holiday Working Days': emp.holidayWorkingDays || 0,
      'Total Present Days (Inc. Sundays)': emp.totalPresentDaysIncSundays || (emp.daysPresent + (emp.holidayWorkingDays || 0)),
      'Avg Login': emp.avgClockIn || '--',
      'Avg Logout': emp.avgClockOut || '--',
      'Hours Active': emp.hoursActive,
      'Monthly Salary': emp.monthlySalary || 0,
      'Salary Acquired': emp.salaryAcquired || 0
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    autoFitColumns(ws, exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Salary Summary");
    XLSX.writeFile(wb, `Monthly_Salary_Report_${salaryMonth}.csv`, { bookType: 'csv' });
  };

  const fetchDailyAttendance = async (dateStr) => {
    setLoadingAttendance(true);
    try {
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/attendance/daily?date=${dateStr}`);
      setDailyAttendanceData(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('Failed to fetch daily attendance', err);
    } finally {
      setLoadingAttendance(false);
    }
  };

  const fetchSalaryReport = async (monthStr) => {
    setLoadingSalary(true);
    try {
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/salary/monthly?month=${monthStr}`);
      setSalaryReport(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('Failed to fetch salary report', err);
    } finally {
      setLoadingSalary(false);
    }
  };

  const fetchHolidays = async (monthStr) => {
    setLoadingCalendar(true);
    try {
      const res = await axios.get(`${import.meta.env.VITE_API_URL}/api/holidays?month=${monthStr}`);
      setHolidays(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('Failed to fetch holidays', err);
    } finally {
      setLoadingCalendar(false);
    }
  };

  const handleToggleHoliday = async (dateStr) => {
    try {
      const res = await axios.post(`${import.meta.env.VITE_API_URL}/api/holidays/toggle`, { date: dateStr, title: 'Public Holiday' });
      if (res.data && res.data.success) {
        fetchHolidays(calendarMonth);
        fetchSalaryReport(salaryMonth);
      }
    } catch (err) {
      console.error('Failed to toggle holiday', err);
      alert('Failed to update public holiday status');
    }
  };

  useEffect(() => {
    if (selectedEmployeeId) {
      fetchDetailedData(selectedEmployeeId, detailedDate);
    }
  }, [selectedEmployeeId, detailedDate, lastUpdate]);

  useEffect(() => {
    if (isAuthenticated) {
      fetchDailyAttendance(attendanceDate);
    }
  }, [attendanceDate, isAuthenticated]);

  useEffect(() => {
    if (isAuthenticated && activeTab === 'salary') {
      fetchSalaryReport(salaryMonth);
    }
  }, [salaryMonth, isAuthenticated, activeTab]);

  useEffect(() => {
    if (isAuthenticated && (activeTab === 'calendar' || activeTab === 'salary')) {
      fetchHolidays(activeTab === 'calendar' ? calendarMonth : salaryMonth);
    }
  }, [calendarMonth, salaryMonth, isAuthenticated, activeTab]);

  useEffect(() => {
    if (!isAuthenticated) return;
    fetchData();
    fetchLiveData();

    // Fetch state data every 30 seconds
    const dataFetchInterval = setInterval(() => {
      fetchLiveData();
      fetchData();
    }, 30000);

    // 1-second UI refresh
    const uiRefreshInterval = setInterval(() => {
      setUiTick(prev => prev + 1);
    }, 1000);

    // Connect WebSocket for real-time updates
    const socket = io(import.meta.env.VITE_API_URL);

    socket.on('data-update', () => {
      fetchData();
    });

    socket.on('live-update', () => {
      // Whenever a desktop agent pings the server, this event fires instantly
      fetchLiveData();
      setLastUpdate(Date.now()); // Triggers the detailed modal to refresh if it's open
    });

    socket.on('app-activity-update', (newAppAct) => {
      setDetailedData(prevData => {
        if (!prevData || prevData.employee.id !== newAppAct.employeeId) return prevData;
        // Prepend to top since it is inverted (newest first)
        return {
          ...prevData,
          appActivities: [newAppAct, ...(prevData.appActivities || [])]
        };
      });
    });

    return () => {
      clearInterval(dataFetchInterval);
      clearInterval(uiRefreshInterval);
      socket.disconnect();
    };
  }, [isAuthenticated]);

  const handleAuth = async (e) => {
    e.preventDefault();
    setAuthError('');
    try {
      await axios.post(`${import.meta.env.VITE_API_URL}/api/admin/login`, authForm);
      setIsAuthenticated(true);
      setLoggedInEmail(authForm.email);
      localStorage.setItem('adminAuth', 'true');
      localStorage.setItem('adminEmail', authForm.email);
    } catch (err) {
      setAuthError(err.response?.data?.error || 'Authentication failed');
    }
  };


  const handleApprove = async (id, status) => {
    setProcessingLeaves(prev => ({ ...prev, [id]: true }));
    try {
      await axios.put(`${import.meta.env.VITE_API_URL}/api/leaves/${id}`, { status });
      fetchData();
    } catch (err) {
      console.error('Error updating status', err);
    } finally {
      setProcessingLeaves(prev => ({ ...prev, [id]: false }));
    }
  };

  const formatDateRange = (leave) => {
    const start = new Date(leave.startDate).toLocaleDateString();
    const end = new Date(leave.endDate).toLocaleDateString();
    let dates = start === end ? start : `${start} to ${end}`;

    if (leave.leaveType === 'HALF_DAY') {
      dates += ` (Half Day - ${leave.duration})`;
    } else if (leave.leaveType === 'HOURLY') {
      dates += ` (Hourly - ${leave.duration} Hours)`;
    }

    return dates;
  };

  const handleAddEmployee = async (e) => {
    e.preventDefault();
    setIsAddingEmployee(true);
    try {
      await axios.post(`${import.meta.env.VITE_API_URL}/api/employees`, newEmployee);
      setNewEmployee({ name: '', email: '', monthlySalary: '' });
      fetchData();
    } catch (err) {
      alert(err.response?.data?.error || 'Error adding employee');
    } finally {
      setIsAddingEmployee(false);
    }
  };

  const handleDeleteEmployee = async (id) => {
    if (!window.confirm("Are you sure you want to delete this employee? This will also delete their leaves.")) return;
    try {
      await axios.delete(`${import.meta.env.VITE_API_URL}/api/employees/${id}`);
      fetchData();
    } catch (err) {
      console.error('Error deleting employee', err);
    }
  };

  const handleEditClick = (emp) => {
    setEditingEmployeeId(emp.id);
    setEditFormData({ newId: emp.id, name: emp.name, email: emp.email, password: emp.password || '', monthlySalary: emp.monthlySalary || '' });
  };

  const handleSaveEdit = async (id) => {
    try {
      const res = await axios.put(`${import.meta.env.VITE_API_URL}/api/employees/${id}`, editFormData);
      setEditingEmployeeId(null);
      fetchData();
    } catch (err) {
      console.error('Error updating employee', err);
      alert(err.response?.data?.error || 'Failed to update employee details');
    }
  };

  const handleDeleteLeave = async (id) => {
    if (!window.confirm("Are you sure you want to delete this leave request?")) return;
    try {
      await axios.delete(`${import.meta.env.VITE_API_URL}/api/leaves/${id}`);
      fetchData();
    } catch (err) {
      console.error('Error deleting leave', err);
    }
  };

  const handlePrevDate = () => {
    const d = new Date(attendanceDate);
    d.setDate(d.getDate() - 1);
    setAttendanceDate(d.toISOString().split('T')[0]);
  };

  const handleNextDate = () => {
    const d = new Date(attendanceDate);
    d.setDate(d.getDate() + 1);
    setAttendanceDate(d.toISOString().split('T')[0]);
  };

  const checkLateLogin = (dateStr) => {
    if (!dateStr) return false;
    const d = new Date(dateStr);
    const minutes = d.getHours() * 60 + d.getMinutes();
    return minutes > (9 * 60 + 15); // After 9:15 AM
  };

  const checkEarlyLogout = (dateStr) => {
    if (!dateStr) return false;
    const d = new Date(dateStr);
    const minutes = d.getHours() * 60 + d.getMinutes();
    return minutes < (18 * 60 + 30); // Before 6:30 PM
  };

  if (!isAuthenticated) {
    return (
      <div className="login-wrapper" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '100vh', fontFamily: "'Inter', sans-serif" }}>
        <div className="glass-panel floating-animation" style={{ width: '400px', margin: '20px' }}>
          <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '30px', background: 'rgba(255,255,255,0.8)', padding: '15px', borderRadius: '12px' }}>
            <img src="/logo.png" alt="Subhada Polymers" style={{ height: '100px', objectFit: 'contain' }} />
          </div>
          <form onSubmit={handleAuth} style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
            <input
              type="email"
              placeholder="Admin Email"
              className="edit-input"
              value={authForm.email}
              onChange={e => setAuthForm({ ...authForm, email: e.target.value })}
              required
            />
            <div style={{ position: 'relative' }}>
              <input
                type={showPassword ? "text" : "password"}
                placeholder="Password"
                className="edit-input"
                value={authForm.password}
                onChange={e => setAuthForm({ ...authForm, password: e.target.value })}
                required
              />
              <span
                onClick={() => setShowPassword(!showPassword)}
                style={{ position: 'absolute', right: '15px', top: '50%', transform: 'translateY(-50%)', cursor: 'pointer', opacity: 0.7, fontSize: '18px' }}
                title={showPassword ? "Hide Password" : "Show Password"}
              >
                {showPassword ? '🙈' : '👁️'}
              </span>
            </div>
            {authError && <div style={{ color: '#ef4444', fontSize: '13px', textAlign: 'center' }}>{authError}</div>}
            <button type="submit" className="btn btn-primary" style={{ padding: '14px', fontSize: '16px', marginTop: '10px' }}>
              Secure Login
            </button>
          </form>
        </div>
        <style>{`
          *, *::before, *::after { box-sizing: border-box; }
          body, html, #root { margin: 0; padding: 0; width: 100%; min-height: 100vh; overflow-x: hidden; }
          .login-wrapper {
            background: radial-gradient(circle at 15% 50%, rgba(25, 118, 210, 0.4), transparent 50%),
                        radial-gradient(circle at 85% 30%, rgba(13, 71, 161, 0.4), transparent 50%),
                        linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            background-size: cover;
          }
          .glass-panel { background: rgba(255, 255, 255, 0.1); backdrop-filter: blur(24px); -webkit-backdrop-filter: blur(24px); border: 1px solid rgba(255, 255, 255, 0.2); border-radius: 20px; padding: 40px; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5); }
          .floating-animation { animation: float 6s ease-in-out infinite; }
          @keyframes float { 0% { transform: translateY(0px); } 50% { transform: translateY(-10px); } 100% { transform: translateY(0px); } }
          .edit-input { width: 100%; background: rgba(255, 255, 255, 0.9); border: 1px solid rgba(255,255,255,0.5); padding: 14px 16px; border-radius: 12px; color: #0f172a; outline: none; font-size: 15px; font-family: inherit; transition: all 0.3s; }
          .edit-input:focus { box-shadow: 0 0 0 4px rgba(56, 189, 248, 0.4); border-color: #38bdf8; background: #ffffff; }
          .btn { border: none; cursor: pointer; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); border-radius: 12px; font-weight: 600; font-family: inherit; }
          .btn-primary { background: linear-gradient(135deg, #38bdf8 0%, #0284c7 100%); color: white; box-shadow: 0 10px 20px -5px rgba(2, 132, 199, 0.5); }
          .btn-primary:hover { transform: translateY(-2px) scale(1.02); box-shadow: 0 15px 25px -5px rgba(2, 132, 199, 0.6); }
        `}</style>
      </div>
    );
  }

  // Compute aggregated app usage
  const appUsageSummary = detailedData && Array.isArray(detailedData.appActivities) ?
    Object.entries(detailedData.appActivities.reduce((acc, act) => {
      acc[act.appName] = (acc[act.appName] || 0) + act.durationSec;
      return acc;
    }, {})).sort((a, b) => b[1] - a[1]).map(([appName, durationSec]) => ({ appName, durationSec }))
    : [];

  return (
    <div className="crm-layout">
      {/* MOBILE MENU BUTTON & OVERLAY */}
      {isMobileMenuOpen && <div className="mobile-menu-overlay" onClick={() => setIsMobileMenuOpen(false)}></div>}

      {/* SIDEBAR */}
      <aside className={`sidebar ${isMobileMenuOpen ? 'open' : ''}`}>
        <div className="sidebar-brand">
          <img src="/logo.png" alt="Subhada Polymers" style={{ width: '100%', maxHeight: '90px', objectFit: 'contain' }} />
        </div>
        <nav className="sidebar-nav">
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('dashboard'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'dashboard' ? 'active' : ''}`}>
            <span className="nav-icon">📊</span> Dashboard
          </a>
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('employees'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'employees' ? 'active' : ''}`}>
            <span className="nav-icon">👥</span> Employees Directory
          </a>
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('leaves'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'leaves' ? 'active' : ''}`}>
            <span className="nav-icon">📅</span> Leave Requests
          </a>
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('attendance'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'attendance' ? 'active' : ''}`}>
            <span className="nav-icon">📋</span> Attendance Log
          </a>
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('salary'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'salary' ? 'active' : ''}`}>
            <span className="nav-icon">💰</span> Salary Summary
          </a>
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('calendar'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'calendar' ? 'active' : ''}`}>
            <span className="nav-icon">🗓️</span> Calendar
          </a>
          <a href="#" onClick={(e) => { e.preventDefault(); setActiveTab('live'); setIsMobileMenuOpen(false); }} className={`nav-item ${activeTab === 'live' ? 'active' : ''}`}>
            <span className="nav-icon">📍</span> Live Tracking
          </a>
        </nav>
        <div className="sidebar-footer" style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
          <div className="admin-profile">
            <div className="admin-avatar">{loggedInEmail ? loggedInEmail[0].toUpperCase() : 'A'}</div>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontSize: '14px', fontWeight: 'bold', color: 'white' }}>System Admin</div>
              <div style={{ fontSize: '12px', color: '#e0f2fe', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '140px' }}>{loggedInEmail}</div>
            </div>
          </div>
          <button className="logout-btn" style={{ width: '100%', justifyContent: 'center' }} onClick={() => {
            setIsAuthenticated(false);
            localStorage.removeItem('adminAuth');
            localStorage.removeItem('adminEmail');
          }}>
            Logout
          </button>
        </div>
      </aside>

      {/* MAIN CONTENT */}
      <main className="main-content">
        <header className="top-header glass-panel">
          <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
            <button className="mobile-hamburger icon-btn" onClick={() => setIsMobileMenuOpen(true)}>
              ☰
            </button>
            <h1 className="page-title" style={{ margin: 0 }}>
              {activeTab === 'employees' ? 'Employees Directory' : activeTab === 'leaves' ? 'Leave Requests' : activeTab === 'calendar' ? 'Company Calendar & Holidays' : activeTab === 'live' ? 'Live Tracking' : activeTab === 'attendance' ? 'Attendance Log' : activeTab === 'salary' ? 'Salary Summary' : 'Overview Dashboard'}
            </h1>
          </div>
          <div className="header-actions">
            <div style={{ position: 'relative', display: 'inline-block' }}>
              <button className="icon-btn" onClick={() => setActiveTab('leaves')}>🔔</button>
              {pendingLeavesCount > 0 && (
                <span className="glow-badge" onClick={() => setActiveTab('leaves')}>
                  {pendingLeavesCount}
                </span>
              )}
            </div>
          </div>
        </header>

        <div className="dashboard-grid fade-in-up">

          {/* DASHBOARD TAB */}
          {activeTab === 'dashboard' && (
            loadingLeaves ? <Loader /> : (
            <div className="glass-panel stats-grid">
              <div className="stat-card">
                <h3>Total Employees</h3>
                <div className="stat-value">{employees.length}</div>
              </div>
              <div className="stat-card">
                <h3>Pending Leaves</h3>
                <div className="stat-value" style={{ color: '#f59e0b' }}>{pendingLeavesCount}</div>
              </div>
            </div>
            )
          )}

          {/* EMPLOYEES TAB */}
          {activeTab === 'employees' && (
            loadingLeaves ? <Loader /> : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
              <div className="glass-panel">
                <h2 className="section-title">Add New Employee</h2>
                <form onSubmit={handleAddEmployee} className="add-emp-form">
                  <input type="text" placeholder="Full Name" value={newEmployee.name} onChange={e => setNewEmployee({ ...newEmployee, name: e.target.value })} required />
                  <input type="email" placeholder="Email Address" value={newEmployee.email} onChange={e => setNewEmployee({ ...newEmployee, email: e.target.value })} required />
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', flex: '1 1 200px' }}>
                    <input 
                      type="number" 
                      placeholder="Monthly Salary (₹)" 
                      value={newEmployee.monthlySalary} 
                      onChange={e => setNewEmployee({ ...newEmployee, monthlySalary: e.target.value })} 
                      min="0" 
                      step="any"
                    />
                    {newEmployee.monthlySalary > 0 && (
                      <div style={{ fontSize: '11px', color: '#10b981', fontWeight: '600', paddingLeft: '4px', display: 'flex', flexDirection: 'column' }}>
                        <span>≈ ₹{(newEmployee.monthlySalary / 24).toFixed(2)} / day (24 days)</span>
                        <span>≈ ₹{((newEmployee.monthlySalary / 24) / 8).toFixed(2)} / hr (8 hrs/day)</span>
                      </div>
                    )}
                  </div>
                  <button type="submit" disabled={isAddingEmployee} className="btn btn-primary" style={{ minWidth: '120px' }}>
                    {isAddingEmployee ? 'Adding...' : 'Add Employee'}
                  </button>
                </form>
              </div>

              <div className="glass-panel">
                <h2 className="section-title">Employee Roster</h2>
                <div className="table-container">
                  <table className="premium-table">
                    <thead>
                      <tr>
                        <th>ID Code</th>
                        <th>Details</th>
                        <th>Salary Details</th>
                        <th style={{ textAlign: 'right' }}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {employees.map(emp => (
                        <tr key={emp.id} className="table-row">
                          <td>
                            {editingEmployeeId === emp.id ? (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <input type="text" className="edit-input" placeholder="ID Code" value={editFormData.newId} onChange={e => setEditFormData({ ...editFormData, newId: e.target.value })} />
                                <input type="text" className="edit-input" placeholder="Password" value={editFormData.password} onChange={e => setEditFormData({ ...editFormData, password: e.target.value })} />
                              </div>
                            ) : (
                              <div>
                                <span className="code-badge">{emp.id}</span>
                                {emp.password && <div style={{ fontSize: '12px', color: '#64748b', marginTop: '8px', fontWeight: '500' }}>Pwd: {emp.password}</div>}
                              </div>
                            )}
                          </td>
                          <td>
                            {editingEmployeeId === emp.id ? (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <input type="text" className="edit-input" placeholder="Name" value={editFormData.name} onChange={e => setEditFormData({ ...editFormData, name: e.target.value })} />
                                <input type="email" className="edit-input" placeholder="Email" value={editFormData.email} onChange={e => setEditFormData({ ...editFormData, email: e.target.value })} />
                              </div>
                            ) : (
                              <div>
                                <div className="emp-name">{emp.name}</div>
                                <div className="emp-email">{emp.email}</div>
                              </div>
                            )}
                          </td>
                          <td>
                            {editingEmployeeId === emp.id ? (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                <input 
                                  type="number" 
                                  className="edit-input" 
                                  placeholder="Monthly Salary (₹)" 
                                  value={editFormData.monthlySalary} 
                                  onChange={e => setEditFormData({ ...editFormData, monthlySalary: e.target.value })} 
                                />
                                {editFormData.monthlySalary > 0 && (
                                  <div style={{ fontSize: '11px', color: '#10b981', fontWeight: '600', display: 'flex', flexDirection: 'column' }}>
                                    <span>≈ ₹{(editFormData.monthlySalary / 24).toFixed(2)}/day</span>
                                    <span>≈ ₹{((editFormData.monthlySalary / 24) / 8).toFixed(2)}/hr</span>
                                  </div>
                                )}
                              </div>
                            ) : (
                              <div>
                                <div style={{ fontWeight: '600', color: '#0f172a', fontSize: '13px' }}>
                                  ₹{Number(emp.monthlySalary || 0).toLocaleString('en-IN')} <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 'normal' }}>/ month</span>
                                </div>
                                <div style={{ fontSize: '12px', color: '#10b981', fontWeight: '600', marginTop: '2px' }}>
                                  ₹{Number(emp.dailySalary || (emp.monthlySalary ? emp.monthlySalary / 24 : 0)).toLocaleString('en-IN')} <span style={{ fontSize: '10px', color: '#64748b', fontWeight: 'normal' }}>/ day (24 days)</span>
                                </div>
                                <div style={{ fontSize: '12px', color: '#8b5cf6', fontWeight: '600', marginTop: '2px' }}>
                                  ₹{Number(emp.hourlySalary || (emp.monthlySalary ? (emp.monthlySalary / 24) / 8 : 0)).toLocaleString('en-IN')} <span style={{ fontSize: '10px', color: '#64748b', fontWeight: 'normal' }}>/ hr (8 hrs/day)</span>
                                </div>
                              </div>
                            )}
                          </td>
                          <td style={{ textAlign: 'right' }}>
                            <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
                              {editingEmployeeId === emp.id ? (
                                <>
                                  <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); handleSaveEdit(emp.id); }} className="btn btn-approve">Save</button>
                                  <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); setEditingEmployeeId(null); }} className="btn btn-neutral">Cancel</button>
                                </>
                              ) : (
                                <>
                                  <button type="button" onClick={() => handleEditClick(emp)} className="btn btn-edit">Edit</button>
                                  <button type="button" onClick={() => handleDeleteEmployee(emp.id)} className="btn btn-delete">Remove</button>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
            )
          )}

          {/* LEAVES TAB */}
          {activeTab === 'leaves' && (
            loadingLeaves ? <Loader /> : (
            <div className="glass-panel">
              <h2 className="section-title">Leave Requests</h2>
              <div className="table-container">
                <table className="premium-table">
                  <thead>
                    <tr>
                      <th>Employee</th>
                      <th>Dates Requested</th>
                      <th>Reason</th>
                      <th>Status</th>
                      <th style={{ textAlign: 'center' }}>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {leaves.slice().reverse().map(leave => (
                      <tr key={leave.id} className="table-row">
                        <td>
                          <div className="emp-name">{leave.employee?.name || 'Unknown'}</div>
                        </td>
                        <td style={{ color: '#475569', fontSize: '13px', fontWeight: '500' }}>{formatDateRange(leave)}</td>
                        <td style={{ maxWidth: '200px' }}>
                          <span style={{ display: 'block', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', color: '#64748b', fontSize: '13px' }} title={leave.reason}>
                            {leave.reason}
                          </span>
                        </td>
                        <td>
                          <span className={`status-badge glow-${leave.status.toLowerCase()}`}>
                            {leave.status}
                          </span>
                        </td>
                        <td>
                          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', justifyContent: 'center' }}>
                            {leave.status === 'PENDING' && (
                              <>
                                <button onClick={() => handleApprove(leave.id, 'APPROVED')} disabled={processingLeaves[leave.id]} className="btn btn-approve">
                                  {processingLeaves[leave.id] ? 'Processing...' : 'Approve'}
                                </button>
                                <button onClick={() => handleApprove(leave.id, 'REJECTED')} disabled={processingLeaves[leave.id]} className="btn btn-reject">
                                  {processingLeaves[leave.id] ? 'Processing...' : 'Reject'}
                                </button>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            )
          )}

          {/* LIVE TRACKING TAB */}
          {activeTab === 'live' && (
            loadingLive ? <Loader /> : (
            <div className="glass-panel">
              <h2 className="section-title">Live Employee Tracking</h2>
              <div className="live-grid">
                {liveData.map(emp => (
                  <div key={emp.id} className="live-card" onClick={() => setSelectedEmployeeId(emp.id)}>
                    <div className="live-card-header">
                      <div className="emp-name">{emp.name}</div>
                      <div className={`pulse-indicator ${emp.status.toLowerCase()}`}></div>
                    </div>
                    <div className="emp-email" style={{ marginBottom: '15px' }}>{emp.department} • {emp.employeeId}</div>

                    <div className="live-card-stats">
                      <div className="stat-box">
                        <span className="stat-label">STATUS</span>
                        <span className={`status-badge ${emp.status.toLowerCase()}`}>
                          {emp.status === 'ADMIN_DECLINED' ? 'Access Declined' : emp.status === 'TEMP_ACTIVE' ? `Temp Active: ${emp.tempReason || ''}` : emp.status}
                        </span>
                      </div>
                      <div className="stat-box">
                        <span className="stat-label">ACTIVE HOURS</span>
                        <strong style={{ color: '#0ea5e9' }}>{Math.floor(emp.totalMinutes / 60)}h {emp.totalMinutes % 60}m</strong>
                      </div>
                    </div>
                  </div>
                ))}
                {liveData.length === 0 && (
                  <div style={{ gridColumn: '1 / -1', textAlign: 'center', padding: '40px', color: '#94a3b8' }}>No active employees found.</div>
                )}
              </div>
            </div>
            )
          )}

          {/* ATTENDANCE LOG TAB */}
          {activeTab === 'attendance' && (
            <div className="glass-panel">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', flexWrap: 'wrap', gap: '15px' }}>
                <h2 className="section-title" style={{ margin: 0 }}>Daily Attendance Ledger</h2>
                <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                  <button className="btn btn-primary" onClick={downloadTodaysReport}>Download Today's Report</button>
                  <button className="btn btn-neutral" onClick={downloadMonthlyReport}>Download Detailed Monthly Report</button>
                  <div className="date-selector" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <input type="month" className="edit-input date-picker-input" value={reportMonth} onChange={(e) => setReportMonth(e.target.value)} onClick={(e) => e.target.showPicker?.()} style={{ width: 'auto' }} />
                    <button className="btn btn-primary" onClick={downloadSimpleMonthlyReport} style={{ background: '#10b981', color: 'white' }}>Download Simple Monthly Report</button>
                  </div>
                  <div className="date-selector">
                    <button className="btn-neutral date-arrow" onClick={handlePrevDate}>&lt;</button>
                    <input type="date" className="edit-input date-picker-input" value={attendanceDate} onChange={(e) => setAttendanceDate(e.target.value)} onClick={(e) => e.target.showPicker?.()} />
                    <button className="btn-neutral date-arrow" onClick={handleNextDate}>&gt;</button>
                  </div>
                </div>
              </div>

              {loadingAttendance ? <Loader /> : (
              <div className="table-container">
                <table className="premium-table">
                  <thead>
                    <tr>
                      <th>Employee</th>
                      <th>Clock In</th>
                      <th>Clock Out</th>
                      <th>Active Hours</th>
                      <th>Final Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dailyAttendanceData.map(emp => {
                      const isLate = checkLateLogin(emp.clockIn);
                      const isEarly = checkEarlyLogout(emp.clockOut);
                      const metRequirement = emp.totalMinutes >= 420; // 7 hours

                      return (
                        <tr key={emp.id} className="table-row">
                          <td>
                            <div className="emp-name">{emp.name}</div>
                            <div className="emp-email">{emp.department} • {emp.employeeId}</div>
                          </td>
                          <td>
                            {emp.clockIn ? (
                              <div>
                                <strong style={{ color: isLate ? '#ef4444' : '#0f172a' }}>{new Date(emp.clockIn).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</strong>
                                {isLate && <span className="penalty-badge">Late Login</span>}
                                {emp.systemBootTime && (
                                  <div style={{ fontSize: '11px', color: '#f59e0b', marginTop: '4px', fontWeight: '500' }}>System Active: {new Date(emp.systemBootTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</div>
                                )}
                                {emp.tempReason && (
                                  <div style={{ fontSize: '11px', color: '#3b82f6', marginTop: '4px', fontWeight: '500' }}>Temp Logout: {emp.tempReason}</div>
                                )}
                              </div>
                            ) : (
                              <span style={{ color: '#94a3b8' }}>--</span>
                            )}
                          </td>
                          <td>
                            {emp.clockOut ? (
                              <div>
                                <strong style={{ color: isEarly ? '#ef4444' : '#0f172a' }}>{new Date(emp.clockOut).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</strong>
                                {isEarly && <span className="penalty-badge">Early Logout</span>}
                              </div>
                            ) : emp.clockIn ? (
                              attendanceDate === new Date().toISOString().split('T')[0] ? (
                                <span style={{ color: '#3b82f6', fontWeight: '500' }}>Still Active</span>
                              ) : (
                                <span style={{ color: '#ef4444', fontWeight: '500' }}>Missing Logout</span>
                              )
                            ) : (
                              <span style={{ color: '#94a3b8' }}>--</span>
                            )}
                          </td>
                          <td>
                            <strong style={{ color: metRequirement ? '#10b981' : emp.totalMinutes >= 60 ? '#f59e0b' : '#ef4444' }}>
                              {Math.floor(emp.totalMinutes / 60)}h {emp.totalMinutes % 60}m
                            </strong>
                          </td>
                          <td>
                            {metRequirement ? (
                              <span className="status-badge active" style={{ display: 'inline-block', width: 'auto' }}>Present</span>
                            ) : emp.totalMinutes >= 60 ? (
                              <span className="status-badge idle" style={{ display: 'inline-block', width: 'auto' }}>Requirement Not Met</span>
                            ) : emp.onLeave ? (
                              <span className="status-badge" style={{ display: 'inline-block', width: 'auto', background: '#e0e7ff', color: '#4f46e5' }}>On Leave ({emp.leaveType === 'HALF_DAY' ? 'Half' : emp.leaveType === 'HOURLY' ? 'Hourly' : 'Full'})</span>
                            ) : (
                              <span className="status-badge offline" style={{ display: 'inline-block', width: 'auto' }}>Absent</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                    {dailyAttendanceData.length === 0 && (
                      <tr><td colSpan="5" style={{ textAlign: 'center', padding: '30px', color: '#94a3b8' }}>No employees found.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              )}
            </div>
          )}

          {/* SALARY TAB */}
          {activeTab === 'salary' && (
            <div className="glass-panel">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px', marginBottom: '20px' }}>
                <h2 className="section-title" style={{ margin: 0 }}>Monthly Salary Report</h2>
                <div style={{ display: 'flex', gap: '12px', alignItems: 'center', flexWrap: 'wrap' }}>
                  <button className="btn btn-primary" onClick={downloadSalaryReport} style={{ background: '#10b981', color: 'white' }}>
                    📥 Download Report
                  </button>
                  <input
                    type="month"
                    className="edit-input"
                    style={{ width: '160px', margin: 0, cursor: 'pointer' }}
                    value={salaryMonth}
                    onChange={(e) => setSalaryMonth(e.target.value)}
                    onClick={(e) => e.target.showPicker?.()}
                  />
                  <input
                    type="text"
                    className="edit-input"
                    style={{ width: '200px', margin: 0 }}
                    placeholder="Search employee..."
                    value={salarySearch}
                    onChange={(e) => setSalarySearch(e.target.value)}
                  />
                </div>
              </div>

              {loadingSalary ? <Loader /> : (
                <div className="table-container">
                  <table className="premium-table">
                    <thead>
                      <tr>
                        <th>1<br/>Employee Name</th>
                        <th style={{ textAlign: 'center' }}>2<br/>Total Days</th>
                        <th style={{ textAlign: 'center' }}>3<br/>Days Present</th>
                        <th style={{ textAlign: 'center' }}>4<br/>Days Absent</th>
                        <th style={{ textAlign: 'center' }}>5<br/>Holiday Working Days</th>
                        <th style={{ textAlign: 'center' }}>3+5 (6)<br/>Total Present Days (Inc. Sundays)</th>
                        <th style={{ textAlign: 'center' }}>7<br/>Avg Login</th>
                        <th style={{ textAlign: 'center' }}>8<br/>Avg Logout</th>
                        <th style={{ textAlign: 'center' }}>9<br/>Hours Active</th>
                        <th style={{ textAlign: 'right' }}>10<br/>Monthly Salary</th>
                        <th style={{ textAlign: 'right' }}>(10*6)/24<br/>Salary Acquired</th>
                      </tr>
                    </thead>
                    <tbody>
                      {salaryReport
                        .filter(item => item.name.toLowerCase().includes(salarySearch.toLowerCase()) || item.email.toLowerCase().includes(salarySearch.toLowerCase()))
                        .map(emp => (
                          <tr key={emp.id} className="table-row">
                            <td>
                              <div className="emp-name">{emp.name}</div>
                              <div className="emp-email" style={{ fontSize: '11px', color: '#64748b' }}>{emp.email}</div>
                            </td>
                            <td style={{ textAlign: 'center', fontWeight: '600' }}>
                              {emp.totalDays}
                            </td>
                            <td style={{ textAlign: 'center' }}>
                              <span className="status-badge glow-approved" style={{ display: 'inline-block', width: 'auto', padding: '4px 10px' }}>
                                {emp.daysPresent}
                              </span>
                            </td>
                            <td style={{ textAlign: 'center' }}>
                              {emp.daysAbsent > 0 ? (
                                <span className="status-badge glow-rejected" style={{ display: 'inline-block', width: 'auto', padding: '4px 10px', background: '#fef2f2', color: '#ef4444' }}>
                                  {emp.daysAbsent}
                                </span>
                              ) : (
                                <span style={{ color: '#94a3b8', fontSize: '13px' }}>0</span>
                              )}
                            </td>
                            <td style={{ textAlign: 'center' }}>
                              {emp.holidayWorkingDays > 0 ? (
                                <span className="status-badge" style={{ display: 'inline-block', width: 'auto', padding: '4px 10px', background: '#f0fdf4', color: '#16a34a', border: '1px solid #bbf7d0', fontWeight: 'bold' }}>
                                  🎉 {emp.holidayWorkingDays}
                                </span>
                              ) : (
                                <span style={{ color: '#94a3b8', fontSize: '13px' }}>0</span>
                              )}
                            </td>
                            <td style={{ textAlign: 'center', fontWeight: 'bold', color: '#0f766e' }}>
                              {emp.totalPresentDaysIncSundays}
                            </td>
                            <td style={{ textAlign: 'center', fontWeight: '600', color: '#0284c7', fontSize: '13px' }}>
                              {emp.avgClockIn || '--'}
                            </td>
                            <td style={{ textAlign: 'center', fontWeight: '600', color: '#4f46e5', fontSize: '13px' }}>
                              {emp.avgClockOut || '--'}
                            </td>
                            <td style={{ textAlign: 'center', fontWeight: '600', color: '#334155' }}>
                              ⏱️ {emp.hoursActive} hrs
                            </td>
                            <td style={{ textAlign: 'right', fontWeight: '600', color: '#475569' }}>
                              ₹{Number(emp.monthlySalary || 0).toLocaleString('en-IN')}
                            </td>
                            <td style={{ textAlign: 'right' }}>
                              <div style={{ fontWeight: 'bold', color: '#10b981', fontSize: '14px' }}>
                                ₹{Number(emp.salaryAcquired || 0).toLocaleString('en-IN')}
                              </div>
                            </td>
                          </tr>
                        ))}
                      {salaryReport.length === 0 && (
                        <tr><td colSpan="11" style={{ textAlign: 'center', padding: '30px', color: '#94a3b8' }}>No salary data available for this month.</td></tr>
                      )}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* CALENDAR TAB */}
          {activeTab === 'calendar' && (
            <div className="glass-panel">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '15px', marginBottom: '24px' }}>
                <div>
                  <h2 className="section-title" style={{ margin: 0 }}>Company Calendar & Public Holidays</h2>
                  <p style={{ margin: '4px 0 0 0', color: '#64748b', fontSize: '13px' }}>
                    Click on any date to mark or unmark it as a Public Holiday. Working on a Public Holiday or Sunday counts towards Holiday Working Days.
                  </p>
                </div>
                <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                  <input
                    type="month"
                    className="edit-input"
                    style={{ width: '160px', margin: 0, cursor: 'pointer' }}
                    value={calendarMonth}
                    onChange={(e) => setCalendarMonth(e.target.value)}
                    onClick={(e) => e.target.showPicker?.()}
                  />
                </div>
              </div>

              {loadingCalendar ? <Loader /> : (() => {
                const [year, month] = calendarMonth.split('-').map(Number);
                const daysInMonthCount = new Date(year, month, 0).getDate();
                const firstDayIndex = new Date(year, month - 1, 1).getDay(); // 0 = Sun
                const todayStr = new Date().toISOString().split('T')[0];

                const dayHeaders = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
                const calendarDays = [];

                for (let i = 0; i < firstDayIndex; i++) {
                  calendarDays.push(null);
                }
                for (let day = 1; day <= daysInMonthCount; day++) {
                  const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                  const isSunday = new Date(year, month - 1, day).getDay() === 0;
                  const holidayObj = holidays.find(h => h.date === dateStr);
                  const isHoliday = !!holidayObj;
                  const isToday = dateStr === todayStr;

                  calendarDays.push({
                    day,
                    dateStr,
                    isSunday,
                    isHoliday,
                    holidayTitle: holidayObj?.title || 'Public Holiday',
                    isToday
                  });
                }

                return (
                  <div className="calendar-container">
                    <div className="calendar-grid-header" style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '10px', marginBottom: '10px', textAlign: 'center', fontWeight: 'bold', color: '#475569' }}>
                      {dayHeaders.map(dh => (
                        <div key={dh} style={{ padding: '8px', color: dh === 'Sun' ? '#ef4444' : '#475569' }}>{dh}</div>
                      ))}
                    </div>
                    <div className="calendar-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '10px' }}>
                      {calendarDays.map((cd, index) => {
                        if (!cd) {
                          return <div key={`empty-${index}`} style={{ minHeight: '95px', background: 'rgba(241, 245, 249, 0.4)', borderRadius: '12px', border: '1px border-dash rgba(203,213,225,0.4)' }}></div>;
                        }
                        return (
                          <div
                            key={cd.dateStr}
                            onClick={() => handleToggleHoliday(cd.dateStr)}
                            title={cd.isHoliday ? 'Click to remove Public Holiday' : 'Click to mark as Public Holiday'}
                            style={{
                              minHeight: '95px',
                              padding: '12px',
                              borderRadius: '14px',
                              cursor: 'pointer',
                              display: 'flex',
                              flexDirection: 'column',
                              justify: 'space-between',
                              transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)',
                              background: cd.isHoliday
                                ? 'linear-gradient(135deg, rgba(254, 226, 226, 0.95) 0%, rgba(254, 202, 202, 0.8) 100%)'
                                : cd.isSunday
                                ? 'linear-gradient(135deg, rgba(240, 253, 244, 0.95) 0%, rgba(220, 252, 231, 0.7) 100%)'
                                : 'rgba(255, 255, 255, 0.85)',
                              border: cd.isToday
                                ? '2px solid #0284c7'
                                : cd.isHoliday
                                ? '1px solid #f87171'
                                : cd.isSunday
                                ? '1px solid #86efac'
                                : '1px solid rgba(203, 213, 225, 0.6)',
                              boxShadow: cd.isToday ? '0 0 12px rgba(2, 132, 199, 0.3)' : '0 4px 6px -1px rgba(0,0,0,0.05)'
                            }}
                          >
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                              <span style={{ fontWeight: 'bold', fontSize: '16px', color: cd.isSunday ? '#16a34a' : cd.isHoliday ? '#dc2626' : '#0f172a' }}>
                                {cd.day}
                              </span>
                              {cd.isToday && (
                                <span style={{ fontSize: '10px', background: '#0284c7', color: 'white', padding: '2px 6px', borderRadius: '8px', fontWeight: 'bold' }}>Today</span>
                              )}
                            </div>

                            <div style={{ marginTop: '6px' }}>
                              {cd.isHoliday && (
                                <span className="status-badge" style={{ display: 'inline-block', fontSize: '10px', background: '#fee2e2', color: '#dc2626', border: '1px solid #fecaca', fontWeight: 'bold', width: '100%', textAlign: 'center', padding: '3px 0' }}>
                                  🚩 Public Holiday
                                </span>
                              )}
                              {cd.isSunday && !cd.isHoliday && (
                                <span style={{ fontSize: '11px', color: '#16a34a', fontWeight: '600', display: 'block', textAlign: 'center' }}>
                                  Sunday
                                </span>
                              )}
                            </div>

                            <div style={{ textAlign: 'right', fontSize: '10px', color: cd.isHoliday ? '#b91c1c' : '#64748b', marginTop: '4px', fontWeight: '500' }}>
                              {cd.isHoliday ? 'Tap to Unmark' : 'Tap to Mark'}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}
            </div>
          )}
        </div>
      </main>

      {/* DETAILED VIEW MODAL */}
      {selectedEmployeeId && (
        <div className="modal-overlay" style={{ padding: '20px', zIndex: 1000 }} onClick={() => setSelectedEmployeeId(null)}>
          <div className="modal-content timeline-modal" style={{ position: 'relative', background: 'transparent', border: 'none', boxShadow: 'none', padding: 0, width: '100%', maxWidth: '95vw', height: '100%', maxHeight: '100%', display: 'flex', flexDirection: 'column' }} onClick={e => e.stopPropagation()}>
            <button className="floating-close-btn" style={{ position: 'absolute', top: '10px', right: '10px', zIndex: 20, margin: 0 }} onClick={() => setSelectedEmployeeId(null)}>
              ✕ Close
            </button>

            {detailedData ? (
              <>
                {/* --- MOBILE LAYOUT (1 Column + Buttons) --- */}
                <div className="mobile-only-layout" style={{ flexDirection: 'column', gap: '20px', flex: 1, overflow: 'hidden' }}>

                  {/* TOP SECTION: Details & Actions */}
                  <div style={{ display: 'flex', gap: '20px', alignItems: 'flex-start' }} className="modal-top-section">

                    {/* LEFT: Details */}
                    <div className="floating-section" style={{ padding: '20px', flex: 1 }}>
                      <div className="modal-header" style={{ marginBottom: '15px' }}>
                        <h2 style={{ margin: 0, fontSize: '18px' }}>{detailedData.employee.name}</h2>
                        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                          <input
                            type="date"
                            className="edit-input"
                            style={{ width: '130px', margin: 0, padding: '5px', cursor: 'pointer' }}
                            value={detailedDate}
                            onChange={(e) => setDetailedDate(e.target.value)}
                            onClick={(e) => e.target.showPicker?.()}
                          />
                          <button 
                            className="btn btn-neutral" 
                            style={{ padding: '5px 10px', height: '32px', display: 'flex', alignItems: 'center' }} 
                            onClick={() => fetchDetailedData(selectedEmployeeId, detailedDate)}
                            title="Refresh Data"
                          >
                            🔄
                          </button>
                        </div>
                      </div>

                      <div className="timeline-summary" style={{ marginBottom: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
                        <div className="summary-box" style={{ flex: 1, padding: '10px' }}>
                          <span>Total Active</span>
                          <strong>{detailedData.attendance ? Math.floor(detailedData.attendance.totalMinutes / 60) + 'h ' + (detailedData.attendance.totalMinutes % 60) + 'm' : '0h 0m'}</strong>
                        </div>
                        <div className="summary-box" style={{ flex: 1, padding: '10px' }}>
                          <span>Clock In</span>
                          <strong>{detailedData.attendance?.clockIn ? new Date(detailedData.attendance.clockIn).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '--'}</strong>
                        </div>
                        <div className="summary-box" style={{ flex: 1, padding: '10px' }}>
                          <span>Clock Out</span>
                          <strong>{detailedData.attendance?.clockOut ? new Date(detailedData.attendance.clockOut).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : (detailedData.attendance?.clockIn ? 'Active' : '--')}</strong>
                        </div>
                      </div>
                    </div>

                    {/* RIGHT: Actions */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', minWidth: '150px', justifyContent: 'center', marginLeft: 'auto', marginTop: '50px' }} className="modal-actions-section">
                      <button className="btn btn-primary" style={{ padding: '12px', background: 'rgba(255,255,255,0.9)', color: '#0f172a', border: '1px solid rgba(0,0,0,0.1)', fontWeight: 'bold' }} onClick={() => setActiveSubModal('appActivity')}>
                        📊 App Activity
                      </button>
                      <button className="btn btn-primary" style={{ padding: '12px', background: 'rgba(255,255,255,0.9)', color: '#0f172a', border: '1px solid rgba(0,0,0,0.1)', fontWeight: 'bold' }} onClick={() => setActiveSubModal('appUsage')}>
                        ⏱️ Total App Usage
                      </button>
                    </div>
                  </div>

                  {/* BOTTOM SECTION: System Activity */}
                  <div className="floating-section" style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                    <h3 style={{ marginTop: 0, marginBottom: '20px' }}>System Activity</h3>
                    <div className="timeline" style={{ overflowY: 'auto', flex: 1, paddingRight: '10px' }}>
                      {detailedData.timeline && detailedData.timeline.length > 0 ? detailedData.timeline.map((act, idx) => (
                        <div className="timeline-item" key={act.id || idx}>
                          <div className="timeline-time">
                            {new Date(act.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                          </div>
                          <div className="timeline-marker">
                            <div className={`marker-dot ${act.status.toLowerCase()}`}></div>
                            {idx !== detailedData.timeline.length - 1 && <div className="marker-line"></div>}
                          </div>
                          <div className="timeline-content" style={{ paddingBottom: '25px' }}>
                            <span style={{ color: '#64748b' }}>Status changed to</span>{' '}
                            <span className={`status-badge ${act.status.toLowerCase()}`} style={{ padding: '4px 8px', fontSize: '11px', verticalAlign: 'middle', marginLeft: '5px' }}>
                              {act.status === 'TEMP_ACTIVE' ? `Temp Active: ${act.tempReason || ''}` : act.status}
                            </span>
                          </div>
                        </div>
                      )) : (
                        <div style={{ color: '#64748b', fontSize: '14px', fontStyle: 'italic' }}>No system activity recorded for this day.</div>
                      )}
                    </div>
                  </div>
                </div>

                {/* --- DESKTOP LAYOUT (3 Columns) --- */}
                <div className="desktop-only-layout modal-columns" style={{ flexDirection: 'row', gap: '20px', flex: 1, overflow: 'hidden' }}>

                  {/* LEFT COLUMN: Details + System Activity */}
                  <div className="left-column" style={{ display: 'flex', flexDirection: 'column', gap: '20px', flex: 1, overflow: 'hidden' }}>
                    {/* SECTION 1: Details */}
                    <div className="floating-section" style={{ padding: '20px' }}>
                      <div className="modal-header" style={{ marginBottom: '15px' }}>
                        <h2 style={{ margin: 0, fontSize: '18px' }}>{detailedData.employee.name}</h2>
                        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                          <input
                            type="date"
                            className="edit-input"
                            style={{ width: '130px', margin: 0, padding: '5px', cursor: 'pointer' }}
                            value={detailedDate}
                            onChange={(e) => setDetailedDate(e.target.value)}
                            onClick={(e) => e.target.showPicker?.()}
                          />
                          <button 
                            className="btn btn-neutral" 
                            style={{ padding: '5px 10px', height: '32px', display: 'flex', alignItems: 'center' }} 
                            onClick={() => fetchDetailedData(selectedEmployeeId, detailedDate)}
                            title="Refresh Data"
                          >
                            🔄
                          </button>
                        </div>
                      </div>

                      <div className="timeline-summary" style={{ marginBottom: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
                        <div className="summary-box" style={{ flex: 1, padding: '10px' }}>
                          <span>Total Active</span>
                          <strong>{detailedData.attendance ? Math.floor(detailedData.attendance.totalMinutes / 60) + 'h ' + (detailedData.attendance.totalMinutes % 60) + 'm' : '0h 0m'}</strong>
                        </div>
                        <div className="summary-box" style={{ flex: 1, padding: '10px' }}>
                          <span>Clock In</span>
                          <strong>{detailedData.attendance?.clockIn ? new Date(detailedData.attendance.clockIn).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '--'}</strong>
                        </div>
                        <div className="summary-box" style={{ flex: 1, padding: '10px' }}>
                          <span>Clock Out</span>
                          <strong>{detailedData.attendance?.clockOut ? new Date(detailedData.attendance.clockOut).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : (detailedData.attendance?.clockIn ? 'Active' : '--')}</strong>
                        </div>
                      </div>
                    </div>

                    {/* SECTION 2: System Activity */}
                    <div className="floating-section" style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                      <h3 style={{ marginTop: 0, marginBottom: '20px' }}>System Activity</h3>
                      <div className="timeline" style={{ overflowY: 'auto', flex: 1, paddingRight: '10px' }}>
                        {detailedData.timeline && detailedData.timeline.length > 0 ? detailedData.timeline.map((act, idx) => (
                          <div className="timeline-item" key={act.id || idx}>
                            <div className="timeline-time">
                              {new Date(act.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                            </div>
                            <div className="timeline-marker">
                              <div className={`marker-dot ${act.status.toLowerCase()}`}></div>
                              {idx !== detailedData.timeline.length - 1 && <div className="marker-line"></div>}
                            </div>
                            <div className="timeline-content" style={{ paddingBottom: '25px' }}>
                              <span style={{ color: '#64748b' }}>Status changed to</span>{' '}
                              <span className={`status-badge ${act.status.toLowerCase()}`} style={{ padding: '4px 8px', fontSize: '11px', verticalAlign: 'middle', marginLeft: '5px' }}>
                                {act.status === 'TEMP_ACTIVE' ? `Temp Active: ${act.tempReason || ''}` : act.status}
                              </span>
                            </div>
                          </div>
                        )) : (
                          <div style={{ color: '#64748b', fontSize: '14px', fontStyle: 'italic' }}>No system activity recorded for this day.</div>
                        )}
                      </div>
                    </div>
                  </div> {/* End of Left Column */}

                  {/* MIDDLE COLUMN: App Activity */}
                  <div className="floating-section" style={{ flex: 1.2, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                    <h3 style={{ marginTop: 0, marginBottom: '20px' }}>App Activity</h3>
                    <div className="timeline" style={{ overflowY: 'auto', flex: 1, paddingRight: '10px' }}>
                      {detailedData.appActivities && detailedData.appActivities.length > 0 ? detailedData.appActivities.map((app, idx) => (
                        <div className="timeline-item" key={app.id || idx}>
                          <div className="timeline-time">
                            {new Date(app.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                          </div>
                          <div className="timeline-marker">
                            <div className="marker-dot active" style={{ background: '#3b82f6', boxShadow: '0 0 0 2px rgba(59, 130, 246, 0.3)' }}></div>
                            {idx !== detailedData.appActivities.length - 1 && <div className="marker-line"></div>}
                          </div>
                          <div className="timeline-content" style={{ paddingBottom: '15px', wordBreak: 'break-word', minWidth: 0 }}>
                            <div style={{ fontWeight: 'bold', color: '#1e293b' }}>{app.appName || 'Unknown App'}</div>
                            <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                              Used for {Math.floor(app.durationSec / 60) > 0 ? `${Math.floor(app.durationSec / 60)}m ` : ''}{app.durationSec % 60}s
                            </div>
                          </div>
                        </div>
                      )) : (
                        <div style={{ color: '#64748b', fontSize: '14px', fontStyle: 'italic' }}>No app activity logged yet.</div>
                      )}
                    </div>
                  </div>

                  {/* RIGHT COLUMN: Total App Usage */}
                  <div className="floating-section" style={{ flex: 1.2, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                    <h3 style={{ marginTop: 0, marginBottom: '20px' }}>Total App Usage</h3>
                    <div className="timeline" style={{ overflowY: 'auto', flex: 1, paddingRight: '10px' }}>
                      {appUsageSummary.length > 0 ? appUsageSummary.map((app, idx) => (
                        <div key={idx} style={{ marginBottom: '15px', paddingBottom: '15px', borderBottom: '1px solid rgba(0,0,0,0.05)', wordBreak: 'break-word', minWidth: 0 }}>
                          <div style={{ fontWeight: 'bold', color: '#1e293b', fontSize: '14px', marginBottom: '6px' }}>{app.appName || 'Unknown App'}</div>
                          <div style={{ fontSize: '13px', color: '#10b981', fontWeight: '600' }}>
                            ⏱️ {Math.floor(app.durationSec / 3600) > 0 ? `${Math.floor(app.durationSec / 3600)}h ` : ''}
                            {Math.floor((app.durationSec % 3600) / 60) > 0 ? `${Math.floor((app.durationSec % 3600) / 60)}m ` : ''}
                            {app.durationSec % 60}s
                          </div>
                        </div>
                      )) : (
                        <div style={{ color: '#64748b', fontSize: '14px', fontStyle: 'italic' }}>No app usage data available.</div>
                      )}
                    </div>
                  </div>

                </div>
              </>
            ) : (
              <div className="floating-section" style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '200px' }}>
                <div style={{ color: '#64748b' }}>Loading timeline...</div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* SUB MODALS */}
      {activeSubModal && (
        <div className="modal-overlay" style={{ padding: '20px', zIndex: 2000 }} onClick={() => setActiveSubModal(null)}>
          <div className="modal-content floating-section" style={{ width: '100%', maxWidth: '500px', height: '80vh', display: 'flex', flexDirection: 'column', position: 'relative' }} onClick={e => e.stopPropagation()}>
            <button className="floating-close-btn" style={{ position: 'absolute', top: '20px', right: '20px', zIndex: 10 }} onClick={() => setActiveSubModal(null)}>
              ✕ Close
            </button>

            {activeSubModal === 'appActivity' && (
              <>
                <h3 style={{ marginTop: 0, marginBottom: '20px' }}>App Activity</h3>
                <div className="timeline" style={{ overflowY: 'auto', flex: 1, paddingRight: '10px' }}>
                  {detailedData?.appActivities && detailedData.appActivities.length > 0 ? detailedData.appActivities.map((app, idx) => (
                    <div className="timeline-item" key={app.id || idx}>
                      <div className="timeline-time">
                        {new Date(app.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                      </div>
                      <div className="timeline-marker">
                        <div className="marker-dot active" style={{ background: '#3b82f6', boxShadow: '0 0 0 2px rgba(59, 130, 246, 0.3)' }}></div>
                        {idx !== detailedData.appActivities.length - 1 && <div className="marker-line"></div>}
                      </div>
                      <div className="timeline-content" style={{ paddingBottom: '15px', wordBreak: 'break-word', minWidth: 0 }}>
                        <div style={{ fontWeight: 'bold', color: '#1e293b' }}>{app.appName || 'Unknown App'}</div>
                        <div style={{ fontSize: '12px', color: '#64748b', marginTop: '4px' }}>
                          Used for {Math.floor(app.durationSec / 60) > 0 ? `${Math.floor(app.durationSec / 60)}m ` : ''}{app.durationSec % 60}s
                        </div>
                      </div>
                    </div>
                  )) : (
                    <div style={{ color: '#64748b', fontSize: '14px', fontStyle: 'italic' }}>No app activity logged yet.</div>
                  )}
                </div>
              </>
            )}

            {activeSubModal === 'appUsage' && (
              <>
                <h3 style={{ marginTop: 0, marginBottom: '20px' }}>Total App Usage</h3>
                <div className="timeline" style={{ overflowY: 'auto', flex: 1, paddingRight: '10px' }}>
                  {appUsageSummary.length > 0 ? appUsageSummary.map((app, idx) => (
                    <div key={idx} style={{ marginBottom: '15px', paddingBottom: '15px', borderBottom: '1px solid rgba(0,0,0,0.05)', wordBreak: 'break-word', minWidth: 0 }}>
                      <div style={{ fontWeight: 'bold', color: '#1e293b', fontSize: '14px', marginBottom: '6px' }}>{app.appName || 'Unknown App'}</div>
                      <div style={{ fontSize: '13px', color: '#10b981', fontWeight: '600' }}>
                        ⏱️ {Math.floor(app.durationSec / 3600) > 0 ? `${Math.floor(app.durationSec / 3600)}h ` : ''}
                        {Math.floor((app.durationSec % 3600) / 60) > 0 ? `${Math.floor((app.durationSec % 3600) / 60)}m ` : ''}
                        {app.durationSec % 60}s
                      </div>
                    </div>
                  )) : (
                    <div style={{ color: '#64748b', fontSize: '14px', fontStyle: 'italic' }}>No app usage data available.</div>
                  )}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      <style>{`
        *, *::before, *::after { box-sizing: border-box; }
        body, html, #root { margin: 0; padding: 0; width: 100%; min-height: 100vh; font-family: 'Inter', sans-serif; color: #0f172a; overflow: hidden; }
        
        /* Ultra-Premium Mesh Background */
        #root {
          background-color: #f1f5f9;
          background-image: 
            radial-gradient(at 0% 0%, rgba(224, 242, 254, 0.8) 0px, transparent 50%),
            radial-gradient(at 100% 0%, rgba(186, 230, 253, 0.8) 0px, transparent 50%),
            radial-gradient(at 100% 100%, rgba(240, 249, 255, 0.8) 0px, transparent 50%),
            radial-gradient(at 0% 100%, rgba(224, 242, 254, 0.8) 0px, transparent 50%);
          background-attachment: fixed;
        }

        .crm-layout { display: flex; height: 100vh; width: 100%; overflow: hidden; padding: 20px; gap: 20px; }

        /* Floating Sidebar Styles */
        .sidebar { 
          width: 280px; 
          background: linear-gradient(160deg, rgba(25, 118, 210, 0.9) 0%, rgba(13, 71, 161, 0.95) 100%); 
          display: flex; flex-direction: column; flex-shrink: 0; 
          border-radius: 24px;
          box-shadow: 0 25px 50px -12px rgba(13, 71, 161, 0.4);
          backdrop-filter: blur(20px);
          -webkit-backdrop-filter: blur(20px);
          border: 1px solid rgba(255, 255, 255, 0.15);
          overflow: hidden;
        }
        .sidebar-brand { padding: 20px; display: flex; align-items: center; justify-content: center; background: rgba(255,255,255,0.9); margin: 20px; border-radius: 16px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
        .sidebar-nav { padding: 10px 20px; flex: 1; }
        .nav-item { display: flex; align-items: center; padding: 14px 18px; color: #e0f2fe; text-decoration: none; border-radius: 12px; margin-bottom: 8px; font-weight: 500; font-size: 15px; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
        .nav-item:hover { background: rgba(255, 255, 255, 0.15); color: white; transform: translateX(5px); }
        .nav-item.active { background: rgba(255, 255, 255, 0.95); color: #0d47a1; font-weight: 600; box-shadow: 0 10px 20px -5px rgba(0,0,0,0.2); transform: scale(1.02); }
        .nav-icon { margin-right: 14px; font-size: 18px; transition: transform 0.3s; }
        .nav-item:hover .nav-icon { transform: scale(1.2); }
        .sidebar-footer { padding: 24px; border-top: 1px solid rgba(255, 255, 255, 0.1); background: rgba(0,0,0,0.1); }
        .admin-profile { display: flex; align-items: center; gap: 14px; }
        .admin-avatar { width: 42px; height: 42px; border-radius: 50%; background: linear-gradient(135deg, #ffffff, #e0f2fe); color: #0d47a1; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 18px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }
        .admin-info { min-width: 0; }
        .admin-info strong { display: block; color: white; font-size: 14px; }
        .admin-info span { display: block; color: #e0f2fe; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px; }
        .logout-btn { margin-top: 15px; width: 100%; padding: 12px; background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2); border-radius: 12px; font-weight: 600; color: #ffffff; cursor: pointer; transition: all 0.2s ease; letter-spacing: 0.5px; }
        .logout-btn:hover { background: rgba(239, 68, 68, 0.9); border-color: transparent; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(239, 68, 68, 0.4); }

        /* Main Content Styles */
        .main-content { flex: 1; display: flex; flex-direction: column; min-width: 0; gap: 20px; border-radius: 24px; }
        
        .top-header { height: 80px; padding: 0 32px; display: flex; align-items: center; justify-content: space-between; border-radius: 20px !important; }
        .page-title { font-size: 22px; font-weight: 700; color: #0f172a; margin: 0; letter-spacing: -0.5px; }
        .header-actions { display: flex; align-items: center; gap: 20px; }
        .icon-btn { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 50%; width: 44px; height: 44px; color: #475569; font-size: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
        .icon-btn:hover { background: #ffffff; color: #0f172a; transform: translateY(-3px) rotate(10deg); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
        .glow-badge { position: absolute; top: -5px; right: -5px; background: linear-gradient(135deg, #ef4444, #dc2626); color: white; border-radius: 50%; width: 22px; height: 22px; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 800; cursor: pointer; box-shadow: 0 4px 10px rgba(239, 68, 68, 0.4); border: 2px solid #ffffff; animation: pulse 2s infinite; }
        @keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7); } 70% { box-shadow: 0 0 0 10px rgba(239, 68, 68, 0); } 100% { box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); } }

        .dashboard-grid { overflow-y: auto; flex: 1; padding-bottom: 20px; padding-right: 5px; }
        .dashboard-grid::-webkit-scrollbar { width: 8px; }
        .dashboard-grid::-webkit-scrollbar-track { background: transparent; }
        .dashboard-grid::-webkit-scrollbar-thumb { background: rgba(148, 163, 184, 0.3); border-radius: 10px; }
        
        .fade-in-up { animation: fadeInUp 0.6s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
        @keyframes fadeInUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }

        /* True Glass Panels */
        .glass-panel { 
          background: rgba(255, 255, 255, 0.6); 
          backdrop-filter: blur(24px); 
          -webkit-backdrop-filter: blur(24px); 
          border: 1px solid rgba(255, 255, 255, 0.8); 
          border-radius: 24px; 
          padding: 32px; 
          box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.05), inset 0 0 0 1px rgba(255,255,255,0.5); 
          transition: transform 0.4s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .glass-panel:hover {
          transform: translateY(-4px);
          box-shadow: 0 30px 50px -15px rgba(0, 0, 0, 0.08), inset 0 0 0 1px rgba(255,255,255,0.6);
        }
        
        .section-title { margin-top: 0; color: #0f172a; font-weight: 700; font-size: 20px; margin-bottom: 24px; letter-spacing: -0.5px; }
        
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 24px; }
        .stat-card { background: rgba(255,255,255,0.5); padding: 24px; border-radius: 16px; border: 1px solid rgba(255,255,255,0.8); text-align: center; }
        .stat-card h3 { margin: 0 0 10px 0; color: #64748b; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; }
        .stat-value { font-size: 42px; font-weight: 800; color: #0f172a; letter-spacing: -1px; }

        /* Tables */
        .table-container { overflow-x: auto; border-radius: 12px; }
        .premium-table { width: 100%; border-collapse: separate; border-spacing: 0 8px; text-align: left; }
        .premium-table th { padding: 16px 20px; font-size: 12px; text-transform: uppercase; letter-spacing: 1.5px; color: #64748b; font-weight: 700; }
        .premium-table td { padding: 16px 20px; vertical-align: middle; background: rgba(255,255,255,0.4); transition: background 0.2s; }
        .premium-table td:first-child { border-top-left-radius: 12px; border-bottom-left-radius: 12px; }
        .premium-table td:last-child { border-top-right-radius: 12px; border-bottom-right-radius: 12px; }
        .table-row:hover td { background: rgba(255,255,255,0.8); }
        
        .emp-name { font-weight: 700; color: #0f172a; font-size: 15px; }
        .emp-email { color: #64748b; font-size: 13px; margin-top: 4px; }
        .code-badge { background: rgba(2, 132, 199, 0.1); padding: 6px 12px; border-radius: 8px; font-family: 'JetBrains Mono', monospace; font-size: 13px; color: #0284c7; font-weight: 700; border: 1px solid rgba(2,132,199,0.2); }
        
        .status-badge { padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 700; letter-spacing: 0.5px; display: inline-block; box-shadow: 0 4px 10px -2px rgba(0,0,0,0.1); }
        .glow-pending { background: linear-gradient(135deg, #fef08a, #facc15); color: #854d0e; }
        .glow-approved { background: linear-gradient(135deg, #bbf7d0, #4ade80); color: #14532d; }
        .glow-rejected { background: linear-gradient(135deg, #fecaca, #f87171); color: #7f1d1d; }
        
        .active { background: rgba(74, 222, 128, 0.2); color: #15803d; border: 1px solid rgba(74, 222, 128, 0.3); }
        .idle { background: rgba(250, 204, 21, 0.2); color: #a16207; border: 1px solid rgba(250, 204, 21, 0.3); }
        .offline { background: rgba(148, 163, 184, 0.2); color: #475569; border: 1px solid rgba(148, 163, 184, 0.3); }
        .locked { background: rgba(168, 85, 247, 0.2); color: #7e22ce; border: 1px solid rgba(168, 85, 247, 0.3); }
        .admin_declined { background: rgba(239, 68, 68, 0.2); color: #b91c1c; border: 1px solid rgba(239, 68, 68, 0.3); }
        
        .pulse-indicator { width: 12px; height: 12px; border-radius: 50%; }
        .pulse-indicator.active { background: #4ade80; box-shadow: 0 0 12px #4ade80; animation: pulse 2s infinite; }
        .pulse-indicator.idle { background: #facc15; box-shadow: 0 0 12px #facc15; }
        .pulse-indicator.offline { background: #94a3b8; }
        .pulse-indicator.locked { background: #c084fc; box-shadow: 0 0 12px #c084fc; }
        .pulse-indicator.admin_declined { background: #ef4444; box-shadow: 0 0 12px #ef4444; }
        
        @keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(74, 222, 128, 0.7); } 70% { box-shadow: 0 0 0 10px rgba(74, 222, 128, 0); } 100% { box-shadow: 0 0 0 0 rgba(74, 222, 128, 0); } }

        .add-emp-form { display: flex; gap: 16px; background: rgba(255,255,255,0.5); padding: 24px; border-radius: 16px; border: 1px solid rgba(255,255,255,0.8); }
        .add-emp-form input, .edit-input { background: rgba(255,255,255,0.8); border: 1px solid rgba(203, 213, 225, 0.8); padding: 12px 16px; border-radius: 10px; color: #0f172a; outline: none; flex: 1; font-size: 14px; font-family: inherit; font-weight: 500; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
        .add-emp-form input:focus, .edit-input:focus { border-color: #38bdf8; box-shadow: 0 0 0 4px rgba(56, 189, 248, 0.2); background: #ffffff; transform: translateY(-1px); }

        .date-selector { display: flex; align-items: center; gap: 8px; background: rgba(255,255,255,0.5); padding: 4px; border-radius: 12px; border: 1px solid rgba(203,213,225,0.5); }
        .date-arrow { padding: 8px 12px !important; font-weight: bold; font-size: 16px; display: flex; align-items: center; justify-content: center; height: 38px; line-height: 1; }
        .date-picker-input { padding: 8px 12px; border: none; background: transparent; height: 38px; font-weight: bold; width: 140px; text-align: center; cursor: pointer; user-select: none; -webkit-user-select: none; }
        input[type="date"], input[type="month"] { user-select: none; -webkit-user-select: none; cursor: pointer; }
        input[type="date"]::-webkit-calendar-picker-indicator, input[type="month"]::-webkit-calendar-picker-indicator { cursor: pointer; }
        .penalty-badge { display: inline-block; background: #fee2e2; color: #dc2626; border: 1px solid #fecaca; border-radius: 4px; padding: 2px 6px; font-size: 10px; font-weight: bold; text-transform: uppercase; margin-left: 8px; vertical-align: middle; }
        
        /* Vibrant Buttons */
        .btn { padding: 12px 24px; border: none; border-radius: 10px; font-weight: 600; cursor: pointer; font-size: 14px; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); font-family: inherit; display: inline-flex; align-items: center; gap: 8px; }
        .btn-primary { background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%); color: #ffffff; box-shadow: 0 10px 15px -3px rgba(34, 197, 94, 0.3); }
        .btn-primary:hover:not(:disabled) { transform: translateY(-3px) scale(1.02); box-shadow: 0 15px 25px -5px rgba(34, 197, 94, 0.4); filter: brightness(1.1); }
        
        .btn-approve { background: rgba(74, 222, 128, 0.15); color: #16a34a; border: 1px solid rgba(74, 222, 128, 0.4); }
        .btn-approve:hover { background: rgba(74, 222, 128, 0.3); color: #14532d; transform: translateY(-2px); }
        
        .btn-reject { background: rgba(248, 113, 113, 0.15); color: #dc2626; border: 1px solid rgba(248, 113, 113, 0.4); }
        .btn-reject:hover { background: rgba(248, 113, 113, 0.3); color: #7f1d1d; transform: translateY(-2px); }
        
        .btn-edit { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; box-shadow: 0 4px 6px rgba(245, 158, 11, 0.2); }
        .btn-edit:hover { transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(245, 158, 11, 0.3); }
        
        .btn-delete { background: rgba(255,255,255,0.8); border: 1px solid rgba(239, 68, 68, 0.3); color: #ef4444; }
        .btn-delete:hover { background: #fee2e2; border-color: #ef4444; transform: translateY(-2px); box-shadow: 0 4px 10px rgba(239, 68, 68, 0.1); }
        
        .btn-neutral { background: rgba(255,255,255,0.8); color: #475569; border: 1px solid rgba(203, 213, 225, 0.8); box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .btn-neutral:hover { background: #ffffff; color: #0f172a; transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05); }

        .modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(15, 23, 42, 0.6); backdrop-filter: blur(8px); z-index: 9999; display: flex; justify-content: center; align-items: center; padding: 20px; }
        .modal-content { background: linear-gradient(135deg, rgba(255,255,255,0.9), rgba(255,255,255,0.95)); border: 1px solid rgba(255,255,255,1); border-radius: 24px; padding: 30px; width: 100%; max-height: 90vh; overflow-y: auto; box-shadow: 0 25px 50px -12px rgba(0,0,0,0.3); position: relative; }
        .floating-section { background: linear-gradient(135deg, rgba(255,255,255,0.9), rgba(255,255,255,0.95)); border: 1px solid rgba(255,255,255,1); border-radius: 20px; padding: 25px; box-shadow: 0 15px 35px -10px rgba(0,0,0,0.15); backdrop-filter: blur(10px); }
        .floating-section::-webkit-scrollbar { width: 6px; }
        .floating-section::-webkit-scrollbar-track { background: rgba(0,0,0,0.05); border-radius: 3px; }
        .floating-section::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.2); border-radius: 3px; }

        .floating-close-btn { position: fixed; top: 30px; right: 30px; background: linear-gradient(135deg, #ef4444, #dc2626); border: none; padding: 12px 24px; border-radius: 30px; display: flex; justify-content: center; align-items: center; font-size: 16px; font-weight: bold; cursor: pointer; color: #ffffff; box-shadow: 0 10px 25px -5px rgba(239, 68, 68, 0.5); z-index: 10000; transition: all 0.3s; }
        .floating-close-btn:hover { transform: translateY(-3px) scale(1.05); box-shadow: 0 15px 30px -5px rgba(239, 68, 68, 0.6); }

        .close-btn { position: absolute; top: 20px; right: 20px; background: rgba(0,0,0,0.05); border: none; width: 32px; height: 32px; border-radius: 50%; display: flex; justify-content: center; align-items: center; font-size: 20px; cursor: pointer; color: #64748b; transition: all 0.2s; }
        .close-btn:hover { background: rgba(0,0,0,0.1); color: #0f172a; transform: rotate(90deg); }

        .modal-content.timeline-modal { max-width: 900px; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; border-bottom: 1px solid rgba(0,0,0,0.1); padding-bottom: 16px; }
        
        .timeline-summary { display: flex; gap: 16px; margin-bottom: 30px; }
        .summary-box { flex: 1; background: rgba(0,0,0,0.03); padding: 16px; border-radius: 12px; text-align: center; display: flex; flex-direction: column; gap: 8px; border: 1px solid rgba(0,0,0,0.05); }
        .summary-box span { font-size: 12px; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }
        .summary-box strong { font-size: 20px; color: #0f172a; }
        
        .timeline-container h3 { font-size: 16px; margin-bottom: 20px; color: #334155; }
        .timeline { display: flex; flex-direction: column; }
        .timeline-item { display: flex; min-height: 60px; }
        .timeline-time { width: 80px; padding-right: 16px; text-align: right; font-size: 13px; color: #64748b; font-weight: 500; padding-top: 2px; }
        .timeline-marker { display: flex; flex-direction: column; align-items: center; width: 24px; margin-right: 16px; }
        .marker-dot { width: 12px; height: 12px; border-radius: 50%; z-index: 2; border: 2px solid #fff; }
        .marker-dot.active { background: #4ade80; box-shadow: 0 0 0 2px rgba(74, 222, 128, 0.3); }
        .marker-dot.idle { background: #facc15; box-shadow: 0 0 0 2px rgba(250, 204, 21, 0.3); }
        .marker-dot.offline { background: #94a3b8; box-shadow: 0 0 0 2px rgba(148, 163, 184, 0.3); }
        .marker-dot.locked { background: #c084fc; box-shadow: 0 0 0 2px rgba(192, 132, 252, 0.3); }
        .marker-line { width: 2px; background: #e2e8f0; flex: 1; margin-top: 4px; margin-bottom: 4px; }
        .timeline-content { flex: 1; padding-bottom: 24px; font-size: 14px; color: #475569; }
        .status-text.active { color: #15803d; font-weight: bold; }
        .status-text.idle { color: #a16207; font-weight: bold; }
        .status-text.offline { color: #475569; font-weight: bold; }
        .status-text.locked { color: #7e22ce; font-weight: bold; }

        .live-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; }
        .live-card { background: rgba(255, 255, 255, 0.6); border: 1px solid rgba(255, 255, 255, 0.8); border-radius: 16px; padding: 20px; cursor: pointer; transition: all 0.3s; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
        .live-card:hover { transform: translateY(-3px); box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); background: rgba(255, 255, 255, 0.9); }
        .live-card-header { display: flex; justify-content: space-between; align-items: center; }
        .live-card-stats { display: flex; gap: 12px; margin-top: 10px; padding-top: 15px; border-top: 1px solid rgba(0,0,0,0.05); }
        .stat-box { display: flex; flex-direction: column; gap: 6px; }
        .stat-label { font-size: 11px; color: #94a3b8; font-weight: 600; text-transform: uppercase; }

        .mobile-hamburger { display: none; }
        .mobile-menu-overlay { display: none; }

        .mobile-only-layout { display: none !important; }
        .desktop-only-layout { display: flex !important; }

        @media (max-width: 768px) {
          .desktop-only-layout { display: none !important; }
          .mobile-only-layout { display: flex !important; }
          
          .crm-layout { flex-direction: column; padding-bottom: 0; }
          .mobile-hamburger { display: flex; font-size: 24px; padding: 5px; }
          .mobile-menu-overlay { display: block; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 999; }
          
          .sidebar { 
            position: fixed; top: 0; left: 0; bottom: 0; 
            width: 280px; height: 100vh; 
            border-radius: 0; border-right: 1px solid rgba(255,255,255,0.2); 
            flex-direction: column; padding: 20px; z-index: 1000; 
            align-items: stretch; justify-content: flex-start; 
            background: rgba(15, 23, 42, 0.98); 
            transform: translateX(-100%); transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
          }
          .sidebar.open { transform: translateX(0); }
          .sidebar-brand, .sidebar-footer { display: flex; }
          .sidebar-nav { flex-direction: column; flex-wrap: nowrap; justify-content: flex-start; overflow-y: auto; overflow-x: hidden; padding-bottom: 0; }
          .nav-item { flex-direction: row; justify-content: flex-start; padding: 12px 20px; font-size: 15px; margin-bottom: 10px; white-space: normal; }
          .nav-icon { margin-right: 12px; font-size: 20px; }
          
          .main-content { padding: 15px; }

          .top-header { flex-direction: row; gap: 15px; text-align: left; justify-content: space-between; }
          .dashboard-grid { grid-template-columns: 1fr; }
          .add-emp-form { flex-direction: column; }
          .live-grid { grid-template-columns: 1fr; }
          .table-container { overflow-x: auto; -webkit-overflow-scrolling: touch; margin-bottom: 20px; margin: 0 -15px; padding: 0 15px; width: calc(100% + 30px); }
          .timeline-summary { flex-direction: column; }
          .premium-table { min-width: 800px; }
          
          .modal-overlay { padding: 10px !important; }
          .modal-content { padding: 15px; width: 100%; max-height: 95vh; }
          .modal-header { flex-direction: column; align-items: flex-start; gap: 10px; }
          .date-selector { flex-wrap: wrap; justify-content: center; }
          .timeline-item { flex-wrap: wrap; }
          .modal-columns { flex-direction: column !important; overflow-y: auto; }
          .floating-close-btn { top: 15px !important; right: 15px !important; padding: 8px 16px; font-size: 14px; position: absolute; }
          .modal-top-section { flex-direction: row !important; align-items: flex-start !important; }
          .modal-actions-section { width: auto; flex-direction: column !important; min-width: 140px; margin-top: 0 !important; align-self: center !important; }
        }
      `}</style>
    </div>
  );
}

export default App;
