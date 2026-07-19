import React, { useCallback, useEffect, useState } from 'react';
import { BrowserRouter as Router, Routes, Route, NavLink } from 'react-router-dom';
import axios from 'axios';
import confetti from 'canvas-confetti';
import * as XLSX from 'xlsx';

const API_BASE_URL = 'https://dreammore-app.onrender.com';
console.log('App starting. API_BASE_URL is set to:', API_BASE_URL);

function App() {
  const [file, setFile] = useState(null);
  const [status, setStatus] = useState('');
  const [loading, setLoading] = useState(false);
  const [sentStudents, setSentStudents] = useState([]);
  const [historyRecords, setHistoryRecords] = useState([]);
  const [dashboardSummary, setDashboardSummary] = useState({ totalRecordsFound: 0, emailsSuccessfullySent: 0, failedSkipped: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressStatus, setProgressStatus] = useState('Ready');
  const [progressVisible, setProgressVisible] = useState(false);
  const [dashboardSearch, setDashboardSearch] = useState('');
  const [historySearch, setHistorySearch] = useState('');
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [excelRows, setExcelRows] = useState([]);
  const [courseOptions, setCourseOptions] = useState([]);
  const [selectedRecipients, setSelectedRecipients] = useState([]);
  const [selectedCourse, setSelectedCourse] = useState('');
  const [isMobileView, setIsMobileView] = useState(() => window.innerWidth <= 900);
  const [selectedHistoryIds, setSelectedHistoryIds] = useState([]);
  const [historyConfirm, setHistoryConfirm] = useState({ open: false, type: null, count: 0, ids: [] });
  const [downloadConfirm, setDownloadConfirm] = useState({ open: false, format: null, generating: false });
  const [dashboardCurrentPage, setDashboardCurrentPage] = useState(1);
  const [historyCurrentPage, setHistoryCurrentPage] = useState(1);

  const DASHBOARD_ROWS_PER_PAGE = 5;
  const HISTORY_ROWS_PER_PAGE = 10;

  const normalize = (value = '') => String(value).trim().toLowerCase();
  const getStudentKey = (student) => `${normalize(student.Name || student.name || '')}|${normalize(student.Email || student.email || '')}`;
  const getHistoryRecordId = (record) => String(record._id || record.id || getStudentKey(record));

  const mapHistoryRecords = (records) => records.map((record) => ({
    _id: record._id || record.id,
    Name: record.studentName || record.Name || record.name || 'N/A',
    Email: record.email || record.Email || '',
    Course: record.course || record.Course || '',
    sentAt: record.sentAt || record.SentDate || null
  }));

  const resetProgress = () => {
    setProgress(0);
    setProgressStatus('Ready');
    setProgressVisible(false);
  };

  const loadLocalState = useCallback(() => {
    try {
      const savedSentStudents = localStorage.getItem('dreammore-sent-students');
      if (savedSentStudents) {
        setSentStudents(JSON.parse(savedSentStudents));
      }

      const savedSummary = localStorage.getItem('dreammore-dashboard-summary');
      if (savedSummary) {
        setDashboardSummary(JSON.parse(savedSummary));
      }
    } catch (error) {
      console.error('Failed to load saved local state:', error);
    }
  }, []);

  const saveLocalState = useCallback(() => {
    try {
      localStorage.setItem('dreammore-sent-students', JSON.stringify(sentStudents));
      localStorage.setItem('dreammore-dashboard-summary', JSON.stringify(dashboardSummary));
    } catch (error) {
      console.error('Failed to save local state:', error);
    }
  }, [sentStudents, dashboardSummary]);

  const fetchHistory = useCallback(async () => {
    try {
      const response = await axios.get(`${API_BASE_URL}/history`);
      if (response.data?.success) {
        const records = mapHistoryRecords(response.data.history || []);
        setHistoryRecords(records);
        setDashboardSummary((prev) => ({
          ...prev,
          totalRecordsFound: records.length,
          emailsSuccessfullySent: records.length
        }));
        setSelectedHistoryIds([]);
        return records;
      }
    } catch (error) {
      console.error('Failed to load history from backend', error);
    }
    return [];
  }, []);

  useEffect(() => {
    resetProgress();
    loadLocalState();
    fetchHistory();
  }, [loadLocalState, fetchHistory]);

  useEffect(() => {
    const handleResize = () => setIsMobileView(window.innerWidth <= 900);
    handleResize();
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  useEffect(() => {
    saveLocalState();
  }, [saveLocalState]);

  useEffect(() => {
    if (typeof EventSource === 'undefined') {
      return undefined;
    }

    const eventSource = new EventSource(`${API_BASE_URL}/progress`);

    eventSource.onmessage = (event) => {
      const data = JSON.parse(event.data);
      const nextPercent = Math.min(Math.max(Number(data.percent) || 0, 0), 99);
      setProgress(nextPercent);
      setProgressStatus(data.status || 'Processing...');
    };

    eventSource.onerror = () => {
      eventSource.close();
    };

    return () => {
      eventSource.close();
    };
  }, []);

  useEffect(() => {
    if (file) {
      resetProgress();
    }
  }, [file]);
  const closeDuplicateModal = () => {
    window.__dreammorePendingDuplicateEvent = null;
    setStatus('');
  };

  const sendEmails = async (e, skipDuplicatePrompt = false) => {
    e.preventDefault();
    if (!file || !selectedCourse) {
      alert('Please choose a file and a course before sending.');
      return;
    }

    const recipientsToSend = excelRows.filter((row) => {
      const matchesCourse = normalize(row.Course) === normalize(selectedCourse);
      const isSelected = selectedRecipients.includes(row.Email);
      return matchesCourse && isSelected && Boolean(row.Email);
    });

    if (!recipientsToSend.length) {
      alert('No recipients are selected for the chosen course.');
      return;
    }

    setLoading(true);
    setProgress(0);
    setProgressStatus('Preparing recipients...');
    setProgressVisible(true);
    setStatus('Preparing and validating recipients...');

    try {
      const validRecipients = recipientsToSend.filter((row) => {
        const email = normalize(row.Email || row.email || '');
        return email.includes('@') && email.includes('.com');
      });
      const invalidRecipientsList = recipientsToSend.filter((row) => {
        const email = normalize(row.Email || row.email || '');
        return !email.includes('@') || !email.includes('.com');
      });

      const existingKeys = new Set(historyRecords.map((entry) => getStudentKey(entry)));
      const duplicates = validRecipients.filter((row) => existingKeys.has(`${normalize(row.Name || row.name || '')}|${normalize(row.Email || row.email || '')}`));

      if (duplicates.length > 0 && !skipDuplicatePrompt) {
        window.__dreammorePendingDuplicateEvent = e;
        setStatus('__DUPLICATE_CONFIRM__');
        setLoading(false);
        setProgress(0);
        setProgressStatus('Ready');
        return;
      }

      const formData = new FormData();
      formData.append('file', file);
      formData.append('courseName', selectedCourse);
      formData.append('selectedRecipients', JSON.stringify(validRecipients.map((row) => row.Email)));

      const response = await axios.post(`${API_BASE_URL}/send-emails`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });

      const currentBatch = response.data.sentTo || [];
      const duplicateCount = response.data.duplicateCount ?? duplicates.length;
      const nextHistory = [...historyRecords, ...currentBatch.map((student) => ({ ...student, Course: student.Course || student.course || selectedCourse }))];
      const uniqueHistory = nextHistory.filter((entry, index, array) => array.findIndex((item) => getStudentKey(item) === getStudentKey(entry)) === index);

      setSentStudents(currentBatch);
      setHistoryRecords(uniqueHistory);
      setStatus(response.data.success ? `Success: ${response.data.message}` : 'No valid recipients were sent.');
      setProgress(100);
      setProgressStatus('Completed');
      setProgressVisible(true);

      if (response.data.success && currentBatch.length) {
        confetti({
          particleCount: 140,
          spread: 90,
          origin: { y: 0.6 },
          colors: ['#2563eb', '#22c55e', '#f59e0b', '#ec4899']
        });
      }

      const refreshedHistory = await fetchHistory();
      const finalHistory = refreshedHistory.length ? refreshedHistory : uniqueHistory;
      setHistoryRecords(finalHistory);
      setDashboardSummary({
        totalRecordsFound: finalHistory.length,
        emailsSuccessfullySent: finalHistory.length,
        failedSkipped: invalidRecipientsList.length + duplicateCount
      });
    } catch (error) {
      console.error(error);
      setProgress(0);
      setProgressStatus('Failed');
      setProgressVisible(true);
      setStatus('An error occurred while processing the batch.');
    } finally {
      setLoading(false);
    }
  };

  const continueDuplicateResend = async () => {
    if (!window.__dreammorePendingDuplicateEvent) return;

    const resendEvent = window.__dreammorePendingDuplicateEvent;
    window.__dreammorePendingDuplicateEvent = null;
    setStatus('Preparing and validating recipients...');
    await sendEmails(resendEvent, true);
  };

  const readFileAsArrayBuffer = (selectedFile) => {
    if (selectedFile?.arrayBuffer) {
      return selectedFile.arrayBuffer();
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error || new Error('Failed to read the selected file.'));
      reader.readAsArrayBuffer(selectedFile);
    });
  };

  const parseExcelFile = useCallback(async (selectedFile) => {
    if (!selectedFile) return;

    try {
      const arrayBuffer = await readFileAsArrayBuffer(selectedFile);
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const normalizedRows = rows
        .filter((row) => row && (row.Name || row.name || row.Email || row.email || row.Course || row.course))
        .map((row) => ({
          Name: row.Name || row.name || 'N/A',
          Email: row.Email || row.email || '',
          Course: row.Course || row.course || '',
          selected: true
        }));

      const uniqueCourses = Array.from(new Set(normalizedRows.map((row) => String(row.Course).trim()).filter(Boolean))).sort();
      setExcelRows(normalizedRows);
      setCourseOptions(uniqueCourses);
      setSelectedRecipients(normalizedRows.map((row) => row.Email));
      setSelectedCourse('');
      setStatus(`Loaded ${normalizedRows.length} student records from ${selectedFile.name}`);
    } catch (error) {
      console.error('Failed to parse excel file', error);
      setStatus('Unable to parse the selected Excel file.');
    }
  }, []);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      resetProgress();
      setStatus(`Selected file: ${selectedFile.name}`);
      parseExcelFile(selectedFile);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile) {
      setFile(droppedFile);
      resetProgress();
      setStatus(`Selected file: ${droppedFile.name}`);
      parseExcelFile(droppedFile);
    }
  };

  const downloadCsv = (rows, filename) => {
    if (!rows.length) return;

    const headers = ['Name', 'Email', 'Course'];
    const csvRows = [headers.join(',')];

    rows.forEach((row) => {
      const values = headers.map((header) => {
        const value = row[header] || '';
        return `"${String(value).replace(/"/g, '""')}"`;
      });
      csvRows.push(values.join(','));
    });

    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  };

  const downloadCurrentReport = () => {
    const rows = visiblePreviewRows.map((student) => ({
      Name: student.Name || student.name || '',
      Email: student.Email || student.email || '',
      Course: student.Course || student.course || selectedCourse
    }));

    if (!rows.length) {
      setStatus('No filtered recipients are available to export.');
      return;
    }

    downloadCsv(rows, 'recipient-list.csv');
  };

  const openDownloadConfirmation = (format) => {
    setDownloadConfirm({ open: true, format, generating: false });
  };

  const getDownloadFormatLabel = (format) => {
    if (format === 'excel') return 'Excel';
    if (format === 'pdf') return 'PDF';
    if (format === 'recipient-list') return 'Recipient List';
    return 'Report';
  };

  const closeDownloadConfirmation = () => {
    if (downloadConfirm.generating) return;
    setDownloadConfirm({ open: false, format: null, generating: false });
  };

  const runDownloadExport = async () => {
    if (!downloadConfirm.format) return;

    setDownloadConfirm((prev) => ({ ...prev, generating: true }));

    try {
      if (downloadConfirm.format === 'excel') {
        downloadHistoryExcel();
      } else if (downloadConfirm.format === 'pdf') {
        await downloadHistoryPDF();
      } else if (downloadConfirm.format === 'recipient-list') {
        downloadCurrentReport();
      }
    } finally {
      setDownloadConfirm({ open: false, format: null, generating: false });
    }
  };

  const buildHistoryExportRows = () => filteredHistory.map((student) => ({
    'Full Name': student.Name || student.name || '',
    Email: student.Email || student.email || '',
    Course: student.Course || student.course || '',
    Date: student.sentAt ? new Date(student.sentAt).toLocaleString() : 'N/A'
  }));

  const downloadHistoryExcel = () => {
    const rows = buildHistoryExportRows();

    if (!rows.length) return;

    const worksheet = XLSX.utils.json_to_sheet(rows, { header: ['Full Name', 'Email', 'Course', 'Date'] });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'History');

    worksheet['!cols'] = [
      { wch: 24 },
      { wch: 32 },
      { wch: 24 },
      { wch: 28 }
    ];

    const fileName = `dreammore-history-report-${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  const imageUrlToDataUrl = async (url) => {
    const response = await fetch(url);
    const blob = await response.blob();

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  };

  const downloadHistoryPDF = async () => {
    const rows = buildHistoryExportRows();

    if (!rows.length) return;

    const [{ default: jsPDF }, autoTableModule] = await Promise.all([
      import('jspdf'),
      import('jspdf-autotable')
    ]);
    const autoTable = autoTableModule.default;

    const pdf = new jsPDF('p', 'mm', 'a4');
    const generationDate = new Date();
    const formattedDate = generationDate.toLocaleDateString();
    const formattedTime = generationDate.toLocaleTimeString();

    try {
      const logoDataUrl = await imageUrlToDataUrl('https://i.postimg.cc/cHdnkM74/photo-2024-10-15-01-22-00.jpg');
      pdf.addImage(logoDataUrl, 'JPEG', 14, 12, 32, 32);
    } catch (error) {
      console.error('Failed to load PDF logo', error);
    }

    pdf.setFont('helvetica', 'bold');
    pdf.setFontSize(18);
    pdf.setTextColor(26, 29, 33);
    pdf.text('Official Enrollment & Communication Report', 52, 24);
    pdf.setFont('helvetica', 'normal');
    pdf.setFontSize(10);
    pdf.setTextColor(75, 85, 99);
    pdf.text('DreamMore Enrollment Portal', 52, 31);

    autoTable(pdf, {
      startY: 48,
      head: [['Full Name', 'Email', 'Course', 'Date']],
      body: rows.map((row) => [row.Name, row.Email, row.Course, row.Date]),
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 9,
        cellPadding: 4,
        textColor: [31, 41, 55],
        lineColor: [226, 232, 240],
        lineWidth: 0.2
      },
      headStyles: {
        fillColor: [255, 126, 51],
        textColor: [255, 255, 255],
        fontStyle: 'bold'
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      margin: { left: 14, right: 14 },
      didDrawPage: (data) => {
        const pageHeight = typeof pdf.internal.pageSize.getHeight === 'function'
          ? pdf.internal.pageSize.getHeight()
          : pdf.internal.pageSize.height;
        pdf.setDrawColor(226, 232, 240);
        pdf.setLineWidth(0.2);
        pdf.line(14, pageHeight - 18, 196, pageHeight - 18);
        pdf.setFontSize(9);
        pdf.setTextColor(100, 116, 139);
        pdf.text(`Generated on ${formattedDate} at ${formattedTime}`, 14, pageHeight - 10);
        pdf.text('Right work at right time', 196, pageHeight - 10, { align: 'right' });
      }
    });

    const fileName = `dreammore-history-report-${generationDate.toISOString().slice(0, 10)}.pdf`;
    pdf.save(fileName);
  };

  const clearHistory = () => {
    setHistoryConfirm({ open: true, type: 'clear', count: 0, ids: [] });
  };

  const openDeleteSelectedConfirmation = () => {
    setHistoryConfirm({ open: true, type: 'delete-selected', count: selectedHistoryIds.length, ids: [...selectedHistoryIds] });
  };

  const closeHistoryConfirm = () => {
    setHistoryConfirm({ open: false, type: null, count: 0, ids: [] });
  };

  const executeHistoryAction = async () => {
    try {
      if (historyConfirm.type === 'clear') {
        await axios.delete(`${API_BASE_URL}/history/clear`);
      } else if (historyConfirm.type === 'delete-selected') {
        await axios.delete(`${API_BASE_URL}/history/delete-selected`, {
          data: { ids: historyConfirm.ids }
        });
      }

      closeHistoryConfirm();
      setSelectedHistoryIds([]);
      await fetchHistory();
      setStatus('History updated successfully.');
    } catch (error) {
      console.error('Failed to update history', error);
      setStatus('Unable to update history right now.');
    }
  };

  const updateRecipientSelection = (email, selected) => {
    setSelectedRecipients((prev) => prev.includes(email) ? prev.filter((value) => value !== email) : [...prev, email]);
  };

  const toggleSelectAll = (checked) => {
    const pageEmails = dashboardPagedRows.map((row) => row.Email);
    setSelectedRecipients((prev) => {
      if (checked) {
        return Array.from(new Set([...prev, ...pageEmails]));
      }

      return prev.filter((email) => !pageEmails.includes(email));
    });
  };

  const toggleHistoryRowSelection = (recordId, checked) => {
    setSelectedHistoryIds((prev) => (
      checked ? Array.from(new Set([...prev, recordId])) : prev.filter((id) => id !== recordId)
    ));
  };

  const visiblePreviewRows = excelRows.filter((row) => {
    const query = dashboardSearch.toLowerCase();
    const name = (row.Name || '').toString().toLowerCase();
    const email = (row.Email || '').toString().toLowerCase();
    const courseName = (row.Course || '').toString().toLowerCase();
    const matchesCourse = !selectedCourse || normalize(row.Course) === normalize(selectedCourse);
    return matchesCourse && (name.includes(query) || email.includes(query) || courseName.includes(query));
  });

  const dashboardTotalPages = Math.max(1, Math.ceil(visiblePreviewRows.length / DASHBOARD_ROWS_PER_PAGE));
  const dashboardPagedRows = visiblePreviewRows.slice((dashboardCurrentPage - 1) * DASHBOARD_ROWS_PER_PAGE, dashboardCurrentPage * DASHBOARD_ROWS_PER_PAGE);

  const allVisibleSelected = dashboardPagedRows.length > 0 && dashboardPagedRows.every((row) => selectedRecipients.includes(row.Email));

  const filteredHistory = historyRecords.filter((student) => {
    const query = historySearch.toLowerCase();
    const name = (student.Name || student.name || '').toString().toLowerCase();
    const email = (student.Email || student.email || '').toString().toLowerCase();
    const courseName = (student.Course || student.course || '').toString().toLowerCase();
    return name.includes(query) || email.includes(query) || courseName.includes(query);
  });

  const historyTotalPages = Math.max(1, Math.ceil(filteredHistory.length / HISTORY_ROWS_PER_PAGE));
  const pagedHistory = filteredHistory.slice((historyCurrentPage - 1) * HISTORY_ROWS_PER_PAGE, historyCurrentPage * HISTORY_ROWS_PER_PAGE);

  const allHistoryVisibleSelected = pagedHistory.length > 0 && pagedHistory.every((student) => selectedHistoryIds.includes(getHistoryRecordId(student)));

  const hasSelectedHistoryRecords = selectedHistoryIds.length > 0;

  const toggleHistorySelectAll = (checked) => {
    const visibleIds = pagedHistory.map((student) => getHistoryRecordId(student));
    setSelectedHistoryIds((prev) => {
      if (checked) {
        return Array.from(new Set([...prev, ...visibleIds]));
      }

      return prev.filter((id) => !visibleIds.includes(id));
    });
  };

  useEffect(() => {
    setDashboardCurrentPage(1);
  }, [dashboardSearch, selectedCourse, excelRows.length]);

  useEffect(() => {
    setHistoryCurrentPage(1);
  }, [historySearch, historyRecords.length]);

  useEffect(() => {
    if (dashboardCurrentPage > dashboardTotalPages) {
      setDashboardCurrentPage(dashboardTotalPages);
    }
  }, [dashboardCurrentPage, dashboardTotalPages]);

  useEffect(() => {
    if (historyCurrentPage > historyTotalPages) {
      setHistoryCurrentPage(historyTotalPages);
    }
  }, [historyCurrentPage, historyTotalPages]);

  const palette = isDarkMode
    ? {
        background: '#0B0E11',
        surface: '#1A1D21',
        text: '#F7F8FA',
        muted: '#94A3B8',
        border: 'rgba(255,255,255,0.12)',
        inputBg: '#12161B',
        accent: '#FF7E33',
        secondary: 'linear-gradient(135deg, #8B5CF6 0%, #3B82F6 100%)',
        shadow: 'rgba(0,0,0,0.35)'
      }
    : {
        background: '#F4F7FB',
        surface: '#FFFFFF',
        text: '#0F172A',
        muted: '#64748B',
        border: 'rgba(15,23,42,0.08)',
        inputBg: '#FFFFFF',
        accent: '#FF7E33',
        secondary: 'linear-gradient(135deg, #8B5CF6 0%, #3B82F6 100%)',
        shadow: 'rgba(15,23,42,0.08)'
      };

  const panelStyle = {
    background: palette.surface,
    border: `1px solid ${palette.border}`,
    borderRadius: '20px',
    boxShadow: `0 20px 50px ${palette.shadow}`,
    backdropFilter: 'blur(18px)'
  };

  const navLinkStyle = ({ isActive }) => ({
    display: 'block',
    width: '100%',
    padding: '14px 16px',
    borderRadius: '14px',
    textDecoration: 'none',
    color: isActive ? '#FFFFFF' : palette.muted,
    background: isActive ? 'rgba(255,126,51,0.16)' : 'transparent',
    fontWeight: isActive ? '700' : '600',
    marginBottom: '10px',
    borderLeft: isActive ? `4px solid ${palette.accent}` : '4px solid transparent',
    boxShadow: isActive ? '0 0 0 1px rgba(139,92,246,0.25), 0 12px 24px rgba(139,92,246,0.16)' : 'none',
    transition: 'all 180ms ease'
  });

  return (
    <Router>
      <div style={{ minHeight: '100vh', padding: '24px', paddingTop: '112px', background: `radial-gradient(circle at top left, rgba(139,92,246,0.12), transparent 35%), ${palette.background}`, color: palette.text, fontFamily: 'Inter, Poppins, Arial, sans-serif' }}>
        <header style={{ position: 'fixed', top: 0, left: 0, right: 0, zIndex: 1100, background: 'rgba(11, 14, 17, 0.9)', color: '#fff', padding: '18px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', boxShadow: '0 12px 30px rgba(0,0,0,0.24)', borderRadius: '0 0 18px 18px', border: '1px solid rgba(255,255,255,0.08)', backdropFilter: 'blur(18px)', margin: '0 24px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
            <a href='/' style={{ display: 'flex', alignItems: 'center', marginRight: '6px' }}>
              <img
                src='https://i.postimg.cc/cHdnkM74/photo-2024-10-15-01-22-00.jpg'
                alt='DreamMore logo'
                style={{ height: '48px', width: '48px', objectFit: 'contain', borderRadius: '12px' }}
              />
            </a>
            <div>
              <div style={{ fontSize: '20px', fontWeight: '700' }}>DreamMore Enrollment Portal</div>
              <div style={{ fontSize: '13px', opacity: 0.82 }}>Centralized email operations</div>
            </div>
          </div>
          <button type='button' onClick={() => setIsDarkMode((value) => !value)} style={{ border: '1px solid rgba(255,255,255,0.16)', background: 'rgba(255,255,255,0.06)', color: '#fff', borderRadius: '999px', padding: '10px 14px', cursor: 'pointer', fontWeight: '700' }}>
            {isDarkMode ? '☀️ Light' : '🌙 Dark'}
          </button>
        </header>

        <div style={{ display: 'flex', gap: '24px', flexWrap: isMobileView ? 'wrap' : 'nowrap', alignItems: 'flex-start' }}>
          <aside style={{ width: isMobileView ? '100%' : '280px', minHeight: isMobileView ? 'auto' : 'calc(100vh - 124px)', height: isMobileView ? 'auto' : 'calc(100vh - 128px)', position: isMobileView ? 'static' : 'fixed', left: isMobileView ? 'auto' : '24px', top: isMobileView ? 'auto' : '112px', overflowY: isMobileView ? 'visible' : 'auto', zIndex: 900, ...panelStyle, padding: '22px', background: isDarkMode ? 'rgba(26, 29, 33, 0.85)' : 'rgba(255,255,255,0.9)' }}>
            <div style={{ fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.16em', color: palette.muted, marginBottom: '18px' }}>Navigation</div>
            <NavLink to='/' end style={navLinkStyle}>Dashboard</NavLink>
            <NavLink to='/history' style={navLinkStyle}>Full History</NavLink>
          </aside>

          <main style={{ flex: 1, minWidth: 0, marginLeft: isMobileView ? 0 : '304px', maxHeight: isMobileView ? 'none' : 'calc(100vh - 136px)', overflowY: isMobileView ? 'visible' : 'auto', paddingBottom: '24px' }}>
            <Routes>
              <Route
                path='/'
                element={
                  <div style={{ ...panelStyle, padding: '26px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '22px', gap: '12px', flexWrap: 'wrap' }}>
                      <div>
                        <div style={{ fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.12em', color: palette.muted }}>Dashboard</div>
                        <h1 style={{ margin: '10px 0 0', fontSize: '28px' }}>Send Enrollment Emails</h1>
                      </div>
                      <div style={{ padding: '10px 16px', borderRadius: '999px', background: 'rgba(255,126,51,0.14)', color: palette.accent, fontWeight: '700', fontSize: '13px', border: '1px solid rgba(255,126,51,0.24)' }}>Live</div>
                    </div>

                    <form onSubmit={sendEmails} style={{ display: 'grid', gap: '18px' }}>
                      <div style={{ display: 'grid', gap: '16px' }}>
                        <label style={{ fontWeight: '600', color: palette.text }}>Upload Excel File</label>
                        <div onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }} onDragLeave={() => setIsDragging(false)} onDrop={handleDrop} style={{ border: isDragging ? '2px dashed #FF7E33' : '2px dashed rgba(255,126,51,0.45)', borderRadius: '18px', padding: '28px', textAlign: 'center', background: isDragging ? 'rgba(255,126,51,0.12)' : palette.inputBg, cursor: 'pointer', boxShadow: isDragging ? '0 0 0 1px rgba(255,126,51,0.28), 0 18px 35px rgba(255,126,51,0.12)' : 'inset 0 0 0 1px rgba(255,255,255,0.03)' }}>
                          <input type='file' accept='.xlsx, .xls' onChange={handleFileChange} style={{ display: 'none' }} id='file-upload' aria-label='Upload Excel File' />
                          <label htmlFor='file-upload' style={{ cursor: 'pointer', display: 'block' }}>
                            <div style={{ fontSize: '32px', marginBottom: '10px' }}>⬆️</div>
                            <div style={{ fontWeight: '700', marginBottom: '4px' }}>Drop spreadsheet here or click to browse</div>
                            {file && <div style={{ marginTop: '10px', color: palette.accent, fontWeight: '600' }}>{file.name}</div>}
                          </label>
                        </div>
                      </div>

                      <div style={{ display: 'grid', gap: '8px' }}>
                        <label style={{ fontWeight: '600', color: palette.text }} htmlFor='course-select'>Course</label>
                        <select id='course-select' value={selectedCourse} onChange={(e) => { setSelectedCourse(e.target.value); resetProgress(); }} style={{ width: '100%', padding: '14px 16px', borderRadius: '14px', border: `1px solid ${palette.border}`, outline: 'none', fontSize: '14px', background: palette.inputBg, color: palette.text }}>
                          <option value=''>Select a course</option>
                          {courseOptions.map((option) => (
                            <option key={option} value={option}>{option}</option>
                          ))}
                        </select>
                      </div>

                      <button type='submit' disabled={loading} style={{ background: loading ? '#64748B' : 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '16px', padding: '16px', fontWeight: '700', fontSize: '15px', cursor: 'pointer', transition: 'transform 180ms ease, box-shadow 180ms ease', boxShadow: '0 12px 24px rgba(255,126,51,0.24)' }}>{loading ? 'Sending...' : 'Send Enrollment Emails'}</button>
                    </form>

                    {status && status !== '__DUPLICATE_CONFIRM__' && <div style={{ marginTop: '22px', padding: '16px', borderRadius: '16px', background: 'rgba(255,126,51,0.12)', border: '1px solid rgba(255,126,51,0.2)', color: palette.accent, fontWeight: '600' }}>{status}</div>}

                    {progressVisible && (
                      <div style={{ marginTop: '24px', background: palette.inputBg, borderRadius: '16px', padding: '18px', border: `1px solid ${palette.border}` }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', fontWeight: '700', color: palette.text }}>
                          <span>Sending Progress</span>
                          <button onClick={resetProgress} style={{ background: 'transparent', border: 'none', color: palette.accent, cursor: 'pointer', fontWeight: '700' }}>Reset</button>
                        </div>
                        <div style={{ height: '12px', width: '100%', backgroundColor: isDarkMode ? '#26313A' : '#E2E8F0', borderRadius: '999px', overflow: 'hidden' }}>
                          <div style={{ width: `${progress}%`, height: '100%', background: palette.secondary, transition: 'width 0.35s ease' }} />
                        </div>
                        <div style={{ marginTop: '10px', display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: palette.muted }}>
                          <span>{progressStatus}</span>
                          <span>{progress}%</span>
                        </div>
                      </div>
                    )}

                    <div style={{ marginTop: '24px', padding: '22px', borderRadius: '20px', background: palette.inputBg, border: `1px solid ${palette.border}` }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', gap: '12px', flexWrap: 'wrap' }}>
                        <div>
                          <div style={{ fontSize: '12px', color: palette.muted, textTransform: 'uppercase', letterSpacing: '0.12em' }}>Dashboard Summary</div>
                          <h2 style={{ margin: '8px 0 0', fontSize: '22px' }}>Engagement Summary</h2>
                        </div>
                        <div style={{ fontSize: '13px', color: palette.muted }}>Based on stored history</div>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px' }}>
                        <div style={{ background: 'linear-gradient(135deg, rgba(255,126,51,0.22), rgba(255,126,51,0.08))', color: palette.text, padding: '20px', borderRadius: '18px', border: '1px solid rgba(255,126,51,0.2)' }}>
                          <div style={{ fontSize: '28px', marginBottom: '10px' }}>📊</div>
                          <div style={{ fontSize: '14px', opacity: 0.9 }}>Records Found</div>
                          <div style={{ fontSize: '28px', fontWeight: '700' }}>{dashboardSummary.totalRecordsFound}</div>
                        </div>
                        <div style={{ background: palette.secondary, color: '#fff', padding: '20px', borderRadius: '18px' }}>
                          <div style={{ fontSize: '28px', marginBottom: '10px' }}>✅</div>
                          <div style={{ fontSize: '14px', opacity: 0.9 }}>Emails Sent</div>
                          <div style={{ fontSize: '28px', fontWeight: '700' }}>{dashboardSummary.emailsSuccessfullySent}</div>
                        </div>
                        <div style={{ background: 'linear-gradient(135deg, rgba(255,126,51,0.18), rgba(255,126,51,0.04))', color: palette.text, padding: '20px', borderRadius: '18px', border: '1px solid rgba(255,126,51,0.2)' }}>
                          <div style={{ fontSize: '28px', marginBottom: '10px' }}>⚠️</div>
                          <div style={{ fontSize: '14px', opacity: 0.9 }}>Failed / Skipped</div>
                          <div style={{ fontSize: '28px', fontWeight: '700' }}>{dashboardSummary.failedSkipped}</div>
                        </div>
                      </div>

                      <div style={{ marginTop: '22px', display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: '14px' }}>
                        <h3 style={{ margin: 0 }}>Current Batch</h3>
                        <button onClick={() => openDownloadConfirmation('recipient-list')} style={{ background: 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '14px', padding: '12px 18px', cursor: 'pointer', fontWeight: '700', boxShadow: '0 10px 20px rgba(255,126,51,0.2)' }}>Export Recipient List</button>
                      </div>

                      <div style={{ marginTop: '16px', position: 'relative' }}>
                        <span style={{ position: 'absolute', left: '14px', top: '50%', transform: 'translateY(-50%)', fontSize: '16px', color: palette.muted }}>🔎</span>
                        <input value={dashboardSearch} onChange={(e) => setDashboardSearch(e.target.value)} placeholder='Search current batch' style={{ width: '100%', padding: '14px 14px 14px 44px', borderRadius: '14px', border: `1px solid ${palette.border}`, outline: 'none', fontSize: '14px', background: palette.inputBg, color: palette.text }} />
                      </div>

                      <div style={{ marginTop: '18px', overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', backgroundColor: 'transparent', color: palette.text }}>
                          <thead>
                            <tr style={{ backgroundColor: 'rgba(255,255,255,0.04)', color: palette.text }}>
                              <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left', width: '48px' }}>
                                <input type='checkbox' checked={allVisibleSelected} onChange={(e) => toggleSelectAll(e.target.checked)} aria-label='Select all recipients' />
                              </th>
                              <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Name</th>
                              <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Email</th>
                              <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Course</th>
                            </tr>
                          </thead>
                          <tbody>
                            {dashboardPagedRows.map((student, index) => {
                              const isChecked = selectedRecipients.includes(student.Email);
                              return (
                                <tr key={`${student.Email}-${index}`} style={{ backgroundColor: index % 2 === 0 ? 'rgba(255,255,255,0.03)' : 'transparent', transition: 'background-color 160ms ease' }}>
                                  <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>
                                    <input type='checkbox' checked={isChecked} onChange={() => updateRecipientSelection(student.Email, !isChecked)} aria-label={`Select ${student.Name}`} />
                                  </td>
                                  <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.Name || 'N/A'}</td>
                                  <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.Email || 'N/A'}</td>
                                  <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.Course || 'N/A'}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>

                      <div style={{ marginTop: '18px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap' }}>
                        <div style={{ color: palette.muted, fontSize: '13px', fontWeight: '600' }}>
                          Page {dashboardCurrentPage} of {dashboardTotalPages}
                        </div>
                        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                          <button
                            type='button'
                            onClick={() => setDashboardCurrentPage((page) => Math.max(1, page - 1))}
                            disabled={dashboardCurrentPage === 1}
                            style={{ background: dashboardCurrentPage === 1 ? '#475569' : 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '12px', padding: '10px 16px', cursor: dashboardCurrentPage === 1 ? 'not-allowed' : 'pointer', fontWeight: '700', opacity: dashboardCurrentPage === 1 ? 0.65 : 1 }}
                          >
                            Previous
                          </button>
                          <button
                            type='button'
                            onClick={() => setDashboardCurrentPage((page) => Math.min(dashboardTotalPages, page + 1))}
                            disabled={dashboardCurrentPage === dashboardTotalPages}
                            style={{ background: dashboardCurrentPage === dashboardTotalPages ? '#475569' : 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '12px', padding: '10px 16px', cursor: dashboardCurrentPage === dashboardTotalPages ? 'not-allowed' : 'pointer', fontWeight: '700', opacity: dashboardCurrentPage === dashboardTotalPages ? 0.65 : 1 }}
                          >
                            Next
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                }
              />
              <Route
                path='/history'
                element={
                  <div style={{ ...panelStyle, padding: '26px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '22px', gap: '12px', flexWrap: 'wrap' }}>
                      <div>
                        <div style={{ fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.12em', color: palette.muted }}>Full History</div>
                        <h1 style={{ margin: '10px 0 0', fontSize: '28px' }}>All Sent Email Records</h1>
                      </div>
                      <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
                        <button onClick={() => openDownloadConfirmation('excel')} style={{ background: 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '14px', padding: '14px 18px', cursor: 'pointer', fontWeight: '800', boxShadow: '0 10px 20px rgba(255,126,51,0.18)' }}>
                          📊 Download Excel
                        </button>
                        <button onClick={() => openDownloadConfirmation('pdf')} style={{ background: 'rgba(255,255,255,0.06)', color: palette.text, border: `1px solid ${palette.border}`, borderRadius: '14px', padding: '14px 18px', cursor: 'pointer', fontWeight: '800', display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
                          <span>📄</span>
                          <span>Download PDF</span>
                        </button>
                        {hasSelectedHistoryRecords && (
                          <button onClick={openDeleteSelectedConfirmation} style={{ background: 'linear-gradient(135deg, #DC2626 0%, #EF4444 100%)', color: '#fff', border: 'none', borderRadius: '14px', padding: '14px 18px', cursor: 'pointer', fontWeight: '800', boxShadow: '0 10px 20px rgba(220,38,38,0.18)' }}>
                            Delete Selected
                          </button>
                        )}
                        <button onClick={clearHistory} style={{ background: 'rgba(255,255,255,0.06)', color: palette.text, border: `1px solid ${palette.border}`, borderRadius: '14px', padding: '14px 18px', cursor: 'pointer', fontWeight: '700' }}>Clear History</button>
                      </div>
                    </div>

                    <div style={{ position: 'relative', marginBottom: '20px' }}>
                      <span style={{ position: 'absolute', left: '14px', top: '50%', transform: 'translateY(-50%)', fontSize: '16px', color: palette.muted }}>🔎</span>
                      <input value={historySearch} onChange={(e) => setHistorySearch(e.target.value)} placeholder='Search history by name, email, or course' style={{ width: '100%', padding: '14px 14px 14px 44px', borderRadius: '14px', border: `1px solid ${palette.border}`, outline: 'none', fontSize: '14px', background: palette.inputBg, color: palette.text }} />
                    </div>

                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: '720px', color: palette.text }}>
                        <thead>
                          <tr style={{ backgroundColor: 'rgba(255,255,255,0.04)', color: palette.text }}>
                            <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left', width: '48px' }}>
                              <input type='checkbox' checked={allHistoryVisibleSelected} onChange={(e) => toggleHistorySelectAll(e.target.checked)} aria-label='Select all history records' />
                            </th>
                            <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Name</th>
                            <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Email</th>
                            <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Course</th>
                            <th style={{ padding: '14px', borderBottom: `1px solid ${palette.border}`, textAlign: 'left' }}>Sent At</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pagedHistory.map((student, index) => (
                            <tr key={getHistoryRecordId(student)} style={{ backgroundColor: index % 2 === 0 ? 'rgba(255,255,255,0.03)' : 'transparent' }}>
                              <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>
                                <input
                                  type='checkbox'
                                  checked={selectedHistoryIds.includes(getHistoryRecordId(student))}
                                  onChange={(e) => toggleHistoryRowSelection(getHistoryRecordId(student), e.target.checked)}
                                  aria-label={`Select ${student.Name || student.name || 'history record'}`}
                                />
                              </td>
                              <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.Name || student.name || 'N/A'}</td>
                              <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.Email || student.email || 'N/A'}</td>
                              <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.Course || student.course || 'N/A'}</td>
                              <td style={{ padding: '14px', borderBottom: `1px solid ${palette.border}` }}>{student.sentAt ? new Date(student.sentAt).toLocaleString() : 'N/A'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div style={{ marginTop: '18px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap' }}>
                      <div style={{ color: palette.muted, fontSize: '13px', fontWeight: '600' }}>
                        Page {historyCurrentPage} of {historyTotalPages}
                      </div>
                      <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                        <button
                          type='button'
                          onClick={() => setHistoryCurrentPage((page) => Math.max(1, page - 1))}
                          disabled={historyCurrentPage === 1}
                          style={{ background: historyCurrentPage === 1 ? '#475569' : 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '12px', padding: '10px 16px', cursor: historyCurrentPage === 1 ? 'not-allowed' : 'pointer', fontWeight: '700', opacity: historyCurrentPage === 1 ? 0.65 : 1 }}
                        >
                          Previous
                        </button>
                        <button
                          type='button'
                          onClick={() => setHistoryCurrentPage((page) => Math.min(historyTotalPages, page + 1))}
                          disabled={historyCurrentPage === historyTotalPages}
                          style={{ background: historyCurrentPage === historyTotalPages ? '#475569' : 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#fff', border: 'none', borderRadius: '12px', padding: '10px 16px', cursor: historyCurrentPage === historyTotalPages ? 'not-allowed' : 'pointer', fontWeight: '700', opacity: historyCurrentPage === historyTotalPages ? 0.65 : 1 }}
                        >
                          Next
                        </button>
                      </div>
                    </div>
                  </div>
                }
              />
            </Routes>
          </main>
        </div>

        {historyConfirm.open && (
          <div
            role='dialog'
            aria-modal='true'
            aria-labelledby='history-confirm-title'
            aria-describedby='history-confirm-message'
            onClick={closeHistoryConfirm}
            style={{
              position: 'fixed',
              inset: 0,
              zIndex: 1900,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
              background: 'rgba(5, 8, 12, 0.58)',
              backdropFilter: 'blur(14px)',
              WebkitBackdropFilter: 'blur(14px)',
              animation: 'dreammoreFadeIn 180ms ease-out'
            }}
          >
            <div
              onClick={(event) => event.stopPropagation()}
              style={{
                width: '100%',
                maxWidth: '540px',
                background: '#1A1D21',
                border: '1px solid rgba(255,126,51,0.35)',
                borderRadius: '24px',
                boxShadow: '0 30px 80px rgba(0,0,0,0.5)',
                padding: '28px',
                color: '#F7F8FA',
                transformOrigin: 'center',
                animation: 'dreammoreZoomIn 220ms ease-out'
              }}
            >
              <div style={{ width: '56px', height: '56px', borderRadius: '18px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(255,126,51,0.12)', color: '#FF7E33', marginBottom: '18px', fontSize: '28px' }}>
                ⚠️
              </div>
              <h2 id='history-confirm-title' style={{ margin: '0 0 12px', fontSize: '24px', fontWeight: '800', color: '#FFFFFF' }}>
                {historyConfirm.type === 'clear' ? 'Clear History' : 'Delete Selected'}
              </h2>
              <p id='history-confirm-message' style={{ margin: '0 0 24px', color: 'rgba(247,248,250,0.82)', lineHeight: 1.6, fontSize: '15px' }}>
                {historyConfirm.type === 'clear'
                  ? 'Are you sure you want to wipe ALL records? This cannot be undone.'
                  : `Are you sure you want to delete the ${historyConfirm.count} selected records?`}
              </p>
              <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end', flexWrap: 'wrap' }}>
                <button type='button' onClick={closeHistoryConfirm} style={{ border: '1px solid rgba(255,255,255,0.12)', background: '#475569', color: '#E5E7EB', borderRadius: '14px', padding: '12px 18px', fontWeight: '700', cursor: 'pointer' }}>
                  Cancel
                </button>
                <button type='button' onClick={executeHistoryAction} style={{ border: 'none', background: 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#FFFFFF', borderRadius: '14px', padding: '12px 18px', fontWeight: '800', cursor: 'pointer', boxShadow: '0 12px 24px rgba(255,126,51,0.24)' }}>
                  Yes, Proceed
                </button>
              </div>
            </div>
          </div>
        )}

        {downloadConfirm.open && (
          <div
            role='dialog'
            aria-modal='true'
            aria-labelledby='download-confirm-title'
            aria-describedby='download-confirm-message'
            onClick={closeDownloadConfirmation}
            style={{
              position: 'fixed',
              inset: 0,
              zIndex: 1950,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
              background: 'rgba(5, 8, 12, 0.58)',
              backdropFilter: 'blur(14px)',
              WebkitBackdropFilter: 'blur(14px)',
              animation: 'dreammoreFadeIn 180ms ease-out'
            }}
          >
            <div
              onClick={(event) => event.stopPropagation()}
              style={{
                width: '100%',
                maxWidth: '520px',
                background: '#1A1D21',
                border: '1px solid rgba(255,126,51,0.35)',
                borderRadius: '24px',
                boxShadow: '0 30px 80px rgba(0,0,0,0.5)',
                padding: '28px',
                color: '#F7F8FA',
                transformOrigin: 'center',
                animation: 'dreammoreZoomIn 220ms ease-out'
              }}
            >
              <div style={{ width: '56px', height: '56px', borderRadius: '18px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(255,126,51,0.12)', color: '#FF7E33', marginBottom: '18px', fontSize: '28px' }}>
                ⬇️
              </div>
              <h2 id='download-confirm-title' style={{ margin: '0 0 12px', fontSize: '24px', fontWeight: '800', color: '#FFFFFF' }}>
                Confirm Download
              </h2>
              <p id='download-confirm-message' style={{ margin: '0 0 14px', color: 'rgba(247,248,250,0.82)', lineHeight: 1.6, fontSize: '15px' }}>
                You are about to generate and download the enrollment report. Would you like to proceed?
              </p>
              <div style={{ marginBottom: '22px', padding: '12px 14px', borderRadius: '14px', background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.08)', color: '#E2E8F0', fontSize: '14px' }}>
                <strong style={{ color: '#FFFFFF' }}>Format:</strong> {getDownloadFormatLabel(downloadConfirm.format)}
              </div>
              <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end', flexWrap: 'wrap' }}>
                <button type='button' onClick={closeDownloadConfirmation} disabled={downloadConfirm.generating} style={{ border: '1px solid rgba(255,255,255,0.12)', background: '#475569', color: '#E5E7EB', borderRadius: '14px', padding: '12px 18px', fontWeight: '700', cursor: downloadConfirm.generating ? 'not-allowed' : 'pointer', opacity: downloadConfirm.generating ? 0.75 : 1 }}>
                  Cancel
                </button>
                <button type='button' onClick={runDownloadExport} disabled={downloadConfirm.generating} style={{ border: 'none', background: 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#FFFFFF', borderRadius: '14px', padding: '12px 18px', fontWeight: '800', cursor: downloadConfirm.generating ? 'wait' : 'pointer', boxShadow: '0 12px 24px rgba(255,126,51,0.24)', minWidth: '160px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}>
                  {downloadConfirm.generating ? (
                    <span style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
                      <span style={{ width: '14px', height: '14px', borderRadius: '999px', border: '2px solid rgba(255,255,255,0.45)', borderTopColor: '#FFFFFF', animation: 'dreammoreSpin 0.8s linear infinite' }} />
                      Generating file...
                    </span>
                  ) : 'Proceed'}
                </button>
              </div>
            </div>
          </div>
        )}

        {status === '__DUPLICATE_CONFIRM__' && (
          <div
            role='dialog'
            aria-modal='true'
            aria-labelledby='duplicate-modal-title'
            aria-describedby='duplicate-modal-message'
            onClick={closeDuplicateModal}
            style={{
              position: 'fixed',
              inset: 0,
              zIndex: 2000,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
              background: 'rgba(5, 8, 12, 0.58)',
              backdropFilter: 'blur(14px)',
              WebkitBackdropFilter: 'blur(14px)',
              animation: 'dreammoreFadeIn 180ms ease-out'
            }}
          >
            <div
              onClick={(event) => event.stopPropagation()}
              style={{
                width: '100%',
                maxWidth: '520px',
                background: '#1A1D21',
                border: '1px solid rgba(255,126,51,0.35)',
                borderRadius: '24px',
                boxShadow: '0 30px 80px rgba(0,0,0,0.5)',
                padding: '28px',
                color: '#F7F8FA',
                transformOrigin: 'center',
                animation: 'dreammoreZoomIn 220ms ease-out'
              }}
            >
              <div style={{ width: '56px', height: '56px', borderRadius: '18px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(255,126,51,0.12)', color: '#FF7E33', marginBottom: '18px', fontSize: '28px' }}>
                ⚠️
              </div>
              <h2 id='duplicate-modal-title' style={{ margin: '0 0 12px', fontSize: '24px', fontWeight: '800', color: '#FFFFFF' }}>
                Duplicate Detection
              </h2>
              <p id='duplicate-modal-message' style={{ margin: '0 0 24px', color: 'rgba(247,248,250,0.82)', lineHeight: 1.6, fontSize: '15px' }}>
                Some students in this list have already been notified about this course. Do you want to resend the email to them?
              </p>
              <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end', flexWrap: 'wrap' }}>
                <button type='button' onClick={closeDuplicateModal} style={{ border: '1px solid rgba(255,255,255,0.12)', background: '#2A2F36', color: '#E5E7EB', borderRadius: '14px', padding: '12px 18px', fontWeight: '700', cursor: 'pointer' }}>
                  Cancel
                </button>
                <button type='button' onClick={continueDuplicateResend} style={{ border: 'none', background: 'linear-gradient(135deg, #FF7E33 0%, #FF9B5F 100%)', color: '#FFFFFF', borderRadius: '14px', padding: '12px 18px', fontWeight: '800', cursor: 'pointer', boxShadow: '0 12px 24px rgba(255,126,51,0.24)' }}>
                  Yes, Resend
                </button>
              </div>
            </div>
          </div>
        )}

        <style>{`\n          @keyframes dreammoreZoomIn {\n            from { opacity: 0; transform: scale(0.92) translateY(10px); }\n            to { opacity: 1; transform: scale(1) translateY(0); }\n          }\n\n          @keyframes dreammoreFadeIn {\n            from { opacity: 0; }\n            to { opacity: 1; }\n          }\n        `}</style>
      </div>
    </Router>
  );
}

export default App;
