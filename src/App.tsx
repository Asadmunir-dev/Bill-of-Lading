import React, { useState, useEffect, useRef } from 'react';
import {
  Cloud,
  CloudUpload,
  Database,
  FileSpreadsheet,
  Printer,
  Plus,
  Trash2,
  ZoomIn,
  ZoomOut,
  Maximize2,
  ClipboardList,
  MoreVertical,
  Check,
  X,
  Search,
  CheckCircle,
  AlertCircle,
  LogOut,
  LogIn,
  User as UserIcon,
  ShieldCheck,
  FileText,
  RotateCcw,
} from 'lucide-react';
import { User, onAuthStateChanged } from 'firebase/auth';
import {
  auth,
  loginWithGoogle,
  logoutUser,
  testFirestoreConnection,
} from './firebase';
import {
  saveBillToFirestore,
  subscribeUserBills,
  deleteBillFromFirestore,
  saveUserProfile,
} from './services/billOfLadingService';
import { BillOfLadingRecord, ContainerItem, NotifyParty } from './types';
import { exportBillOfLadingToExcel } from './utils/excelExport';

const ALL_AVAILABLE_CLAUSES = [
  'CLEAN SHIPPED ON BOARD',
  'MARINE BILL OF LADING',
  '21 FREE DAYS AT DESTINATION.',
  '14 FREE DAYS AT DESTINATION.',
  'FREIGHT PREPAID',
  'CONTAINERIZED CARGO (FCL/FCL)',
  'REEFER TEMPERATURE TO BE MAINTAINED AT -18°C',
  "SHIPPER'S LOAD, STOW, WEIGHT AND COUNT",
];

const INITIAL_NOTIFY_PARTIES: NotifyParty[] = [
  {
    id: 'np-1',
    name: 'HUNAN ER-KANG(HK)COMPANY LIMITED',
    address:
      'ADD: ROOM 703 KOWLOON BUILDING 555 NATHAN RD, KL.\nCERTIFICATE NO. : 61603449-000-06-23-7\nCONTACT: CECILIA TEL: +86 15873056089',
  },
  {
    id: 'np-2',
    name: 'NOTIFY PARTY #2 NAME / AGENT',
    address:
      'LAEM CHABANG LOGISTICS TERMINAL 2, CHONBURI 20230, THAILAND\nTEL: +66 38 490 000',
  },
];

const INITIAL_CONTAINERS: ContainerItem[] = [
  {
    id: 'c-1',
    containerNo: 'OSLU2834268',
    netWeightKg: 26000,
    grossWeightKg: 26070.2,
    packageCount: 520,
  },
];

export default function App() {
  // Authentication & Cloud State
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [cloudConnected, setCloudConnected] = useState(false);
  const [isSavingCloud, setIsSavingCloud] = useState(false);
  const [savedBills, setSavedBills] = useState<BillOfLadingRecord[]>([]);
  const [showProfileMenu, setShowProfileMenu] = useState(false);
  const [showSavedModal, setShowSavedModal] = useState(false);
  const [showBulkPasteModal, setShowBulkPasteModal] = useState(false);
  const [showClausesPopover, setShowClausesPopover] = useState(false);
  const [bulkPasteText, setBulkPasteText] = useState('');
  const [bulkReplaceExisting, setBulkReplaceExisting] = useState(true);
  const [savedSearchQuery, setSavedSearchQuery] = useState('');
  const [customClauseInput, setCustomClauseInput] = useState('');

  // Dialog & Prompt State (replaces window.confirm/alert)
  const [showSignInPromptModal, setShowSignInPromptModal] = useState(false);
  const [pendingCloudSave, setPendingCloudSave] = useState(false);
  const [deleteConfirmBill, setDeleteConfirmBill] = useState<{ id: string; blNumber: string } | null>(null);
  const [showNewBillConfirm, setShowNewBillConfirm] = useState(false);

  // Toast Notification State
  const [toastMessage, setToastMessage] = useState<{
    text: string;
    type: 'success' | 'info' | 'error';
  } | null>(null);
  const toastTimerRef = useRef<NodeJS.Timeout | null>(null);

  // Zoom State
  const [zoomLevel, setZoomLevel] = useState(0.9);

  // Bill of Lading Form State
  const [headerTitle, setHeaderTitle] = useState('FARAN SUGARS MILLS LTD');
  const [blNumber, setBlNumber] = useState('OSL-2024-PK40156');
  const [bookingRef, setBookingRef] = useState('OSL-40156/24(CRO-3476/24)');
  const [carrierName, setCarrierName] = useState('MOL / OCEAN SHIPPING LINES');
  const [shipperText, setShipperText] = useState(
    'FARAN SUGAR MILLS LTD\n43-1-E,(B) P.E.C.H.S, BLOCK-6, OFF RAZI ROAD,\nSHAHRAH-E-FAISAL\nKARACHI 75400 PAKISTAN'
  );
  const [referenceNo, setReferenceNo] = useState('REF-FSM-1040MT');
  const [consigneeText, setConsigneeText] = useState('TO ORDER');
  const [freightPayableAt, setFreightPayableAt] = useState('FREIGHT PREPAID');
  const [movement, setMovement] = useState('FCL/FCL');
  const [formENumber, setFormENumber] = useState('AHB-EXP-127261-06122024');
  const [vesselAndVoy, setVesselAndVoy] = useState('MOL PRESENCE   VOYAGE: 19');
  const [placeOfDelivery, setPlaceOfDelivery] = useState('LAEM CHABANG THAILAND');
  const [portOfLoading, setPortOfLoading] = useState('KARACHI, PAKISTAN');
  const [portOfDischarge, setPortOfDischarge] = useState('LAEM CHABANG THAILAND');
  const [marksAndNumbers, setMarksAndNumbers] = useState(
    'FARAN SUGAR MILLS LTD\nSHAIKH BHIRKIO\nGROSS WEIGHT 50.135 KGS\nNET WEIGHT 50 KGS\nEXPIRY DATE DECEMBER 2026'
  );
  const [cargoDescription, setCargoDescription] = useState(
    "40X20' FCL CONTAINING 1040.00 METRIC TONS\nPAKISTAN WHITE REFINED SUGAR\nWITH MINIMUM POLARIZATION 99.70 DEGREES,\nMAXIMUM MOISTURE 0.06 PERCENT ,\nMAXIMUM ASH 0.05 PERCENT, MAXIMUM COLOUR 45 ICUMSA UNITS,\nPACKING: IN 50 KGS NET EACH NEW POLYPROPYLENE BAG WITH INNER LINER\nCROP YEAR 2023-24 AND EXPIRY DATE DECEMBER 2026\nFI NO AHB-EXP-127261-06122024\nNET WEIGHT : 1040.000 MT\nGROSS WEIGHT : 1042.808 M.T\nIRREVOCABLE DOCUMENTARY CREDIT NUMBER ILC34970481124SG DATE:14-11-2024\nFREIGHT PREPAID"
  );
  const [grossWeight, setGrossWeight] = useState('1,042.8080');
  const [grossWeightUnit, setGrossWeightUnit] = useState('M/TONS');
  const [nettWeight, setNettWeight] = useState('1040.000');
  const [nettWeightUnit, setNettWeightUnit] = useState('M/TONS');
  const [totalPackagesCount, setTotalPackagesCount] = useState('20800.00');
  const [totalPackagesUnit, setTotalPackagesUnit] = useState('BAGS');
  const [measurement, setMeasurement] = useState('CBM');
  const [containerSummaryLine, setContainerSummaryLine] = useState(
    '40X20 CONT 520-BAGS NET WT 26000-KGS GR WT 26070.20 KGS EACH CONT'
  );
  const [totalContainersPackages, setTotalContainersPackages] = useState("1X20' FCL");
  const [numberOfOriginals, setNumberOfOriginals] = useState('(3) THREE');
  const [placeAndDateOfIssue, setPlaceAndDateOfIssue] = useState('KARACHI ......');

  const [notifyParties, setNotifyParties] = useState<NotifyParty[]>(INITIAL_NOTIFY_PARTIES);
  const [containers, setContainers] = useState<ContainerItem[]>(INITIAL_CONTAINERS);
  const [specialClauses, setSpecialClauses] = useState<string[]>([
    "SHIPPER'S LOAD, STOW, WEIGHT AND COUNT",
  ]);

  // Show Toast Helper
  const showToast = (text: string, type: 'success' | 'info' | 'error' = 'success') => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToastMessage({ text, type });
    toastTimerRef.current = setTimeout(() => {
      setToastMessage(null);
    }, 4000);
  };

  // Test Cloud connection on boot
  useEffect(() => {
    testFirestoreConnection().then((connected) => {
      setCloudConnected(connected);
    });
  }, []);

  // Listen to Auth State Changes
  useEffect(() => {
    const unsubscribeAuth = onAuthStateChanged(auth, async (user) => {
      setCurrentUser(user);
      setAuthLoading(false);

      if (user) {
        showToast(`Welcome back, ${user.displayName || user.email}! Cloud synced.`, 'success');
        // Save user profile metadata
        saveUserProfile(user.uid, {
          email: user.email || '',
          displayName: user.displayName || '',
          photoURL: user.photoURL || '',
        }).catch(console.error);
      }
    });

    return () => unsubscribeAuth();
  }, []);

  // Real-time Firestore subscription to saved Bills of Lading when user is authenticated
  useEffect(() => {
    if (!currentUser) {
      setSavedBills([]);
      return;
    }

    const unsubscribe = subscribeUserBills(
      currentUser.uid,
      (bills) => {
        setSavedBills(bills);
      },
      (error) => {
        console.error('Error syncing saved bills:', error);
      }
    );

    return () => unsubscribe();
  }, [currentUser]);

  // Handle Google Login
  const handleGoogleLogin = async () => {
    try {
      setAuthLoading(true);
      const user = await loginWithGoogle();
      if (!user) {
        showToast('Google Sign-In popup was closed or cancelled.', 'info');
        return;
      }
      setShowProfileMenu(false);
      setShowSignInPromptModal(false);
      showToast(`Signed in as ${user.email}. Google Cloud connected!`, 'success');

      if (pendingCloudSave) {
        setPendingCloudSave(false);
        await performCloudSave(user.uid);
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      if (msg === 'POPUP_BLOCKED') {
        showToast('Popup was blocked by your browser. Please allow popups.', 'error');
      } else {
        showToast('Google Sign-In was cancelled or could not be completed.', 'error');
      }
    } finally {
      setAuthLoading(false);
    }
  };

  // Handle Google Logout
  const handleGoogleLogout = async () => {
    try {
      await logoutUser();
      setShowProfileMenu(false);
      showToast('Logged out. Your saved Bills of Lading remain safely in Google Cloud.', 'info');
    } catch {
      showToast('Sign out error. Please try again.', 'error');
    }
  };

  // Current Bill of Lading Document payload
  const getCurrentDocumentPayload = (): BillOfLadingRecord => {
    return {
      id: blNumber || `BL-${Date.now()}`,
      userId: currentUser?.uid || 'guest',
      userEmail: currentUser?.email || '',
      headerTitle,
      blNumber,
      bookingRef,
      carrierName,
      shipperText,
      referenceNo,
      consigneeText,
      freightPayableAt,
      formENumber,
      vesselAndVoy,
      placeOfDelivery,
      portOfLoading,
      portOfDischarge,
      marksAndNumbers,
      cargoDescription,
      grossWeight,
      grossWeightUnit,
      nettWeight,
      nettWeightUnit,
      totalPackagesCount,
      totalPackagesUnit,
      measurement,
      containerSummaryLine,
      totalContainersPackages,
      movement,
      numberOfOriginals,
      placeAndDateOfIssue,
      specialClauses: [...specialClauses],
      notifyParties: JSON.parse(JSON.stringify(notifyParties)),
      containers: JSON.parse(JSON.stringify(containers)),
      savedAt: new Date().toLocaleString(),
    };
  };

  // Save to Google Cloud Firestore
  const handleSaveToCloud = async () => {
    if (!currentUser) {
      setPendingCloudSave(true);
      setShowSignInPromptModal(true);
      return;
    }

    await performCloudSave(currentUser.uid);
  };

  const performCloudSave = async (uid: string) => {
    setIsSavingCloud(true);
    try {
      const payload = getCurrentDocumentPayload();
      payload.userId = uid;
      payload.userEmail = currentUser?.email || '';

      await saveBillToFirestore(uid, payload);
      showToast(`Bill of Lading (${blNumber}) saved to Google Cloud Firestore!`, 'success');
    } catch {
      showToast('Failed to save to Google Cloud. Please check connection.', 'error');
    } finally {
      setIsSavingCloud(false);
    }
  };

  // Load a Saved Bill of Lading into the workspace
  const handleLoadSavedBill = (record: BillOfLadingRecord) => {
    setHeaderTitle(record.headerTitle || '');
    setBlNumber(record.blNumber || '');
    setBookingRef(record.bookingRef || '');
    setCarrierName(record.carrierName || '');
    setShipperText(record.shipperText || '');
    setReferenceNo(record.referenceNo || '');
    setConsigneeText(record.consigneeText || '');
    setFreightPayableAt(record.freightPayableAt || '');
    setMovement(record.movement || 'FCL/FCL');
    setFormENumber(record.formENumber || '');
    setVesselAndVoy(record.vesselAndVoy || '');
    setPlaceOfDelivery(record.placeOfDelivery || '');
    setPortOfLoading(record.portOfLoading || '');
    setPortOfDischarge(record.portOfDischarge || '');
    setMarksAndNumbers(record.marksAndNumbers || '');
    setCargoDescription(record.cargoDescription || '');
    setGrossWeight(record.grossWeight || '');
    setGrossWeightUnit(record.grossWeightUnit || 'M/TONS');
    setNettWeight(record.nettWeight || '');
    setNettWeightUnit(record.nettWeightUnit || 'M/TONS');
    setTotalPackagesCount(record.totalPackagesCount || '');
    setTotalPackagesUnit(record.totalPackagesUnit || 'BAGS');
    setMeasurement(record.measurement || 'CBM');
    setContainerSummaryLine(record.containerSummaryLine || '');
    setTotalContainersPackages(record.totalContainersPackages || '');
    setNumberOfOriginals(record.numberOfOriginals || '(3) THREE');
    setPlaceAndDateOfIssue(record.placeAndDateOfIssue || 'KARACHI ......');

    if (record.notifyParties && record.notifyParties.length > 0) {
      setNotifyParties(JSON.parse(JSON.stringify(record.notifyParties)));
    }
    if (record.containers && record.containers.length > 0) {
      setContainers(JSON.parse(JSON.stringify(record.containers)));
    }
    if (record.specialClauses && record.specialClauses.length > 0) {
      setSpecialClauses([...record.specialClauses]);
    }

    setShowSavedModal(false);
    showToast(`Loaded B/L: ${record.blNumber} from Google Cloud!`, 'success');
  };

  // Delete a saved B/L from cloud (via modal)
  const handleDeleteSavedBill = (e: React.MouseEvent, billId: string, bNum: string) => {
    e.stopPropagation();
    setDeleteConfirmBill({ id: billId, blNumber: bNum });
  };

  const executeDeleteSavedBill = async () => {
    if (!currentUser || !deleteConfirmBill) return;
    const target = deleteConfirmBill;
    setDeleteConfirmBill(null);
    try {
      await deleteBillFromFirestore(currentUser.uid, target.id);
      showToast(`Deleted "${target.blNumber}" from Google Cloud.`, 'info');
    } catch {
      showToast('Failed to delete document from cloud.', 'error');
    }
  };

  // Create New Blank B/L (via modal)
  const handleCreateNewBill = () => {
    setShowNewBillConfirm(true);
  };

  const executeCreateNewBill = () => {
    setShowNewBillConfirm(false);
    const randomSuffix = Math.floor(10000 + Math.random() * 90000);
    setBlNumber(`OSL-2024-PK${randomSuffix}`);
    setBookingRef(`OSL-${randomSuffix}/24(CRO-${Math.floor(1000 + Math.random() * 9000)}/24)`);
    setContainers([
      { id: `c-${Date.now()}`, containerNo: '', netWeightKg: 0, grossWeightKg: 0, packageCount: 0 },
    ]);
    setShowSavedModal(false);
    showToast('New blank Bill of Lading created.', 'info');
  };

  // Synchronize company name in header with first line of shipper
  const handleHeaderTitleChange = (newTitle: string) => {
    setHeaderTitle(newTitle);
    const lines = shipperText.split('\n');
    lines[0] = newTitle;
    setShipperText(lines.join('\n'));
  };

  // Synchronize package units across container table and summary line
  const handlePackageUnitChange = (newUnit: string) => {
    setTotalPackagesUnit(newUnit);
    const upper = newUnit.toUpperCase();
    const updatedSummary = containerSummaryLine.replace(
      /(\d+)-(BAGS|PKGS|CARTONS|BOXES|PALLETS|UNITS)/gi,
      `$1-${upper}`
    );
    setContainerSummaryLine(updatedSummary);
  };

  // Notify parties operations
  const handleAddNotifyParty = () => {
    const newParty: NotifyParty = {
      id: `np-${Date.now()}`,
      name: `NOTIFY PARTY #${notifyParties.length + 1}`,
      address: 'FULL ADDRESS & CONTACT DETAILS HERE',
    };
    setNotifyParties([...notifyParties, newParty]);
  };

  const handleRemoveNotifyParty = (id: string) => {
    if (notifyParties.length <= 1) return;
    setNotifyParties(notifyParties.filter((p) => p.id !== id));
  };

  const handleUpdateNotifyParty = (id: string, field: 'name' | 'address', val: string) => {
    setNotifyParties(
      notifyParties.map((p) => (p.id === id ? { ...p, [field]: val } : p))
    );
  };

  // Container inputs & keyboard navigation
  const handleContainerNoChange = (idx: number, val: string) => {
    const upper = val.toUpperCase();
    const updated = [...containers];
    updated[idx].containerNo = upper;

    if (upper.trim() !== '') {
      if (!updated[idx].netWeightKg || updated[idx].netWeightKg === 0) updated[idx].netWeightKg = 26000;
      if (!updated[idx].grossWeightKg || updated[idx].grossWeightKg === 0) updated[idx].grossWeightKg = 26070.2;
      if (!updated[idx].packageCount || updated[idx].packageCount === 0) updated[idx].packageCount = 520;
    } else {
      updated[idx].netWeightKg = 0;
      updated[idx].grossWeightKg = 0;
      updated[idx].packageCount = 0;
    }
    setContainers(updated);
  };

  const handleContainerWeightChange = (
    idx: number,
    field: 'netWeightKg' | 'grossWeightKg' | 'packageCount',
    valStr: string
  ) => {
    const clean = valStr.replace(/,/g, '');
    const num = parseFloat(clean) || 0;
    const updated = [...containers];
    updated[idx][field] = num;
    setContainers(updated);
  };

  const handleContainerKeyDown = (idx: number, e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' || e.key === 'ArrowDown') {
      e.preventDefault();
      if (idx === containers.length - 1) {
        handleAddContainer();
      } else {
        const nextInput = document.getElementById(`cont-no-${idx + 1}`);
        if (nextInput) (nextInput as HTMLInputElement).focus();
      }
    } else if (e.key === 'ArrowUp' && idx > 0) {
      e.preventDefault();
      const prevInput = document.getElementById(`cont-no-${idx - 1}`);
      if (prevInput) (prevInput as HTMLInputElement).focus();
    }
  };

  const handleContainerPaste = (idx: number, e: React.ClipboardEvent<HTMLInputElement>) => {
    const pasteText = e.clipboardData.getData('text');
    if (!pasteText) return;
    const items = pasteText
      .split(/[\r\n\t,]+/)
      .map((s) => s.trim().toUpperCase())
      .filter((s) => s.length > 0);

    if (items.length > 1) {
      e.preventDefault();
      const updated = [...containers];
      for (let i = 0; i < items.length; i++) {
        const targetIdx = idx + i;
        const cData: ContainerItem = {
          id: `c-${Date.now()}-${targetIdx}`,
          containerNo: items[i],
          netWeightKg: 26000,
          grossWeightKg: 26070.2,
          packageCount: 520,
        };
        if (targetIdx < updated.length) {
          updated[targetIdx] = {
            ...updated[targetIdx],
            containerNo: items[i],
            netWeightKg: updated[targetIdx].netWeightKg || 26000,
            grossWeightKg: updated[targetIdx].grossWeightKg || 26070.2,
            packageCount: updated[targetIdx].packageCount || 520,
          };
        } else {
          updated.push(cData);
        }
      }
      setContainers(updated);
    }
  };

  const handleAddContainer = () => {
    setContainers((prev) => [
      ...prev,
      {
        id: `cont-${Date.now()}`,
        containerNo: '',
        netWeightKg: 0,
        grossWeightKg: 0,
        packageCount: 0,
      },
    ]);
  };

  const handleRemoveContainer = (id: string) => {
    if (containers.length <= 1) return;
    setContainers(containers.filter((c) => c.id !== id));
  };

  // Bulk Import
  const handleApplyBulkContainers = () => {
    const items = bulkPasteText
      .split(/[\r\n\t,]+/)
      .map((s) => s.trim().toUpperCase())
      .filter((s) => s.length > 0);

    if (items.length === 0) {
      setShowBulkPasteModal(false);
      return;
    }

    const mapped: ContainerItem[] = items.map((num, i) => ({
      id: `c-${Date.now()}-${i}`,
      containerNo: num,
      netWeightKg: 26000,
      grossWeightKg: 26070.2,
      packageCount: 520,
    }));

    if (bulkReplaceExisting) {
      setContainers(mapped);
    } else {
      setContainers((prev) => [...prev, ...mapped]);
    }

    setShowBulkPasteModal(false);
    setBulkPasteText('');
    showToast(`Imported ${items.length} containers with auto-calculated weights.`, 'success');
  };

  // Clauses selection
  const handleToggleClause = (clause: string) => {
    if (specialClauses.includes(clause)) {
      setSpecialClauses(specialClauses.filter((c) => c !== clause));
    } else {
      setSpecialClauses([...specialClauses, clause]);
    }
  };

  const handleApplyPreset = (preset: 'sugar' | 'reefer' | 'general') => {
    if (preset === 'sugar') {
      setSpecialClauses([
        'CLEAN SHIPPED ON BOARD',
        'MARINE BILL OF LADING',
        '21 FREE DAYS AT DESTINATION.',
        'FREIGHT PREPAID',
        "SHIPPER'S LOAD, STOW, WEIGHT AND COUNT",
      ]);
    } else if (preset === 'reefer') {
      setSpecialClauses([
        'CLEAN SHIPPED ON BOARD',
        'MARINE BILL OF LADING',
        '21 FREE DAYS AT DESTINATION.',
        'FREIGHT PREPAID',
        'CONTAINERIZED CARGO (FCL/FCL)',
        'REEFER TEMPERATURE TO BE MAINTAINED AT -18°C',
        "SHIPPER'S LOAD, STOW, WEIGHT AND COUNT",
      ]);
    } else {
      setSpecialClauses([
        'CLEAN SHIPPED ON BOARD',
        'MARINE BILL OF LADING',
        '14 FREE DAYS AT DESTINATION.',
        'FREIGHT PREPAID',
        'CONTAINERIZED CARGO (FCL/FCL)',
        "SHIPPER'S LOAD, STOW, WEIGHT AND COUNT",
      ]);
    }
    setShowClausesPopover(false);
  };

  const handleAddCustomClause = () => {
    const val = customClauseInput.trim().toUpperCase();
    if (!val) return;
    if (!specialClauses.includes(val)) {
      setSpecialClauses([...specialClauses, val]);
    }
    setCustomClauseInput('');
  };

  // Export to Excel
  const handleExportExcel = async () => {
    try {
      const payload = getCurrentDocumentPayload();
      await exportBillOfLadingToExcel(payload);
      showToast('Excel spreadsheet exported successfully (.xlsx)', 'success');
    } catch (err) {
      console.error('Excel export error:', err);
      showToast('Failed to export Excel file.', 'error');
    }
  };

  // Calculations
  const filledContainers = containers.filter((c) => c.containerNo.trim() !== '');
  const containerCount = filledContainers.length;

  let totalNetKg = 0;
  let totalGrossKg = 0;
  let totalBags = 0;

  filledContainers.forEach((c) => {
    totalNetKg += c.netWeightKg > 0 ? c.netWeightKg : 26000;
    totalGrossKg += c.grossWeightKg > 0 ? c.grossWeightKg : 26070.2;
    totalBags += c.packageCount > 0 ? c.packageCount : 520;
  });

  // Filter saved bills for modal
  const filteredSavedBills = savedBills.filter((item) => {
    const q = savedSearchQuery.toLowerCase();
    return (
      item.blNumber?.toLowerCase().includes(q) ||
      item.headerTitle?.toLowerCase().includes(q) ||
      item.carrierName?.toLowerCase().includes(q) ||
      item.portOfDischarge?.toLowerCase().includes(q)
    );
  });

  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-900 pb-16">
      {/* =========================================================================
          TOP NAVIGATION HEADER
      ========================================================================== */}
      <header className="no-print bg-[#1F4E79] text-white sticky top-0 z-50 shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-2 flex flex-wrap items-center justify-between gap-3">
          {/* Logo & System Brand */}
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded bg-white text-[#1F4E79] flex items-center justify-center font-black text-xs shadow-xs tracking-tight">
              BL
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-sm font-bold tracking-tight text-white">
                  Ocean Shipping Lines — Bill of Lading
                </h1>
                <span className="text-[10px] bg-blue-950 text-blue-200 font-mono px-1.5 py-0.5 rounded border border-blue-400/30 flex items-center gap-1">
                  <Cloud className="w-3 h-3 text-cyan-400" />
                  Google Cloud Firestore
                </span>
              </div>
              <p className="text-[11px] text-blue-200 flex items-center gap-1.5">
                <span>Enterprise Maritime Documentation</span>
                <span>&bull;</span>
                <span className="text-blue-100 flex items-center gap-1">
                  <span className={`w-1.5 h-1.5 rounded-full ${currentUser ? 'bg-emerald-400' : 'bg-amber-400'}`} />
                  {currentUser ? `Saved in Google Account (${currentUser.email})` : 'Offline / Guest Mode'}
                </span>
              </p>
            </div>
          </div>

          {/* Action Bar & Google Profile Integration */}
          <div className="flex items-center flex-wrap gap-2">
            {/* GOOGLE PROFILE & AUTH BUTTON */}
            {currentUser ? (
              <div className="relative">
                <button
                  type="button"
                  onClick={() => setShowProfileMenu(!showProfileMenu)}
                  className="flex items-center gap-2 px-2.5 py-1.5 bg-blue-900/80 hover:bg-blue-800 text-white rounded-lg border border-blue-400/40 text-xs font-semibold cursor-pointer transition-all shadow-sm"
                  title="Google Account Profile & Cloud Storage Settings"
                >
                  {currentUser.photoURL ? (
                    <img
                      src={currentUser.photoURL}
                      alt="Avatar"
                      className="w-5 h-5 rounded-full border border-white/60 object-cover"
                    />
                  ) : (
                    <div className="w-5 h-5 rounded-full bg-blue-500 text-white flex items-center justify-center text-[10px] font-bold">
                      {currentUser.displayName ? currentUser.displayName.charAt(0).toUpperCase() : 'G'}
                    </div>
                  )}
                  <span className="max-w-[120px] truncate font-medium text-blue-100">
                    {currentUser.displayName || currentUser.email}
                  </span>
                  <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" title="Google Cloud Active" />
                </button>

                {/* Profile Dropdown */}
                {showProfileMenu && (
                  <div className="absolute right-0 mt-2 w-72 bg-white rounded-xl shadow-2xl border border-slate-200 p-3.5 z-50 text-slate-800">
                    <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
                      {currentUser.photoURL ? (
                        <img
                          src={currentUser.photoURL}
                          alt="Avatar"
                          className="w-10 h-10 rounded-full border border-slate-200"
                        />
                      ) : (
                        <div className="w-10 h-10 rounded-full bg-blue-600 text-white flex items-center justify-center font-bold text-sm">
                          {currentUser.displayName ? currentUser.displayName.charAt(0) : 'U'}
                        </div>
                      )}
                      <div className="overflow-hidden">
                        <p className="font-bold text-xs text-slate-900 truncate">
                          {currentUser.displayName || 'Google User'}
                        </p>
                        <p className="text-[11px] text-slate-500 truncate">{currentUser.email}</p>
                        <span className="inline-flex items-center gap-1 text-[10px] text-emerald-700 bg-emerald-50 px-1.5 py-0.2 rounded font-semibold mt-1">
                          <ShieldCheck className="w-3 h-3 text-emerald-600" /> Google Cloud Storage Active
                        </span>
                      </div>
                    </div>

                    <div className="py-2.5 text-xs text-slate-600 space-y-1.5">
                      <div className="flex items-center justify-between">
                        <span>Saved Bills in Cloud:</span>
                        <span className="font-mono font-bold text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded">
                          {savedBills.length} documents
                        </span>
                      </div>
                      <p className="text-[11px] text-slate-400 leading-snug">
                        Your documents are permanently stored under your Google profile. When you log out and log back in, all data is automatically restored.
                      </p>
                    </div>

                    <div className="pt-2 border-t border-slate-100 flex items-center justify-between gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setShowProfileMenu(false);
                          setShowSavedModal(true);
                        }}
                        className="text-xs text-blue-700 hover:text-blue-900 font-semibold cursor-pointer"
                      >
                        View Saved List
                      </button>
                      <button
                        type="button"
                        onClick={handleGoogleLogout}
                        className="flex items-center gap-1 px-2.5 py-1 text-xs text-rose-600 hover:bg-rose-50 rounded font-semibold cursor-pointer transition-colors"
                      >
                        <LogOut className="w-3.5 h-3.5" />
                        Sign Out
                      </button>
                    </div>
                  </div>
                )}
              </div>
            ) : (
              <button
                type="button"
                onClick={handleGoogleLogin}
                disabled={authLoading}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-white text-slate-800 hover:bg-slate-50 text-xs font-bold rounded shadow cursor-pointer transition-all border border-slate-300"
                title="Sign in with Google to persistently save and restore your documents in Google Cloud"
              >
                {/* Google Colorful "G" Icon */}
                <svg className="w-3.5 h-3.5" viewBox="0 0 24 24">
                  <path
                    fill="#4285F4"
                    d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"
                  />
                  <path
                    fill="#34A853"
                    d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
                  />
                  <path
                    fill="#FBBC05"
                    d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z"
                  />
                  <path
                    fill="#EA4335"
                    d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z"
                  />
                </svg>
                <span>Google Login</span>
              </button>
            )}

            {/* SAVE TO CLOUD BUTTON */}
            <button
              type="button"
              onClick={handleSaveToCloud}
              disabled={isSavingCloud}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold rounded shadow cursor-pointer transition-colors"
              title="Save this Bill of Lading permanently in Google Cloud Firestore"
            >
              {isSavingCloud ? (
                <div className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
              ) : (
                <CloudUpload className="w-4 h-4" />
              )}
              <span>{isSavingCloud ? 'Saving...' : 'Save to Cloud'}</span>
            </button>

            {/* SAVED LIST MODAL BUTTON */}
            <button
              type="button"
              onClick={() => setShowSavedModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-sky-100 hover:bg-sky-200 text-sky-950 border border-sky-300 text-xs font-semibold rounded shadow cursor-pointer transition-colors"
              title="View all saved Bills of Lading in Google Cloud"
            >
              <Database className="w-4 h-4 text-sky-700" />
              <span>Saved List</span>
              <span className="bg-sky-800 text-white text-[10px] px-1.5 py-0.2 rounded-full font-mono font-bold">
                {savedBills.length}
              </span>
            </button>

            {/* EXCEL EXPORT BUTTON */}
            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded shadow cursor-pointer transition-colors"
              title="Export complete formatted document to Excel .xlsx"
            >
              <FileSpreadsheet className="w-4 h-4" />
              <span>Export (.xlsx)</span>
            </button>

            {/* PRINT A4 BUTTON */}
            <button
              type="button"
              onClick={() => window.print()}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-white text-xs font-semibold rounded shadow cursor-pointer transition-colors"
              title="Print strict A4 portrait page"
            >
              <Printer className="w-4 h-4" />
              <span>Print A4</span>
            </button>
          </div>
        </div>
      </header>

      {/* =========================================================================
          CONTROLS TOOLBAR: QUICK ACTIONS & ZOOM
      ========================================================================== */}
      <div className="no-print bg-white border-b border-slate-200 px-4 py-2">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center flex-wrap gap-2">
            <button
              type="button"
              onClick={handleAddNotifyParty}
              className="inline-flex items-center gap-1 px-2.5 py-1 bg-blue-50 hover:bg-blue-100 text-[#1F4E79] font-bold rounded border border-blue-200 text-xs cursor-pointer transition-colors"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>+ Add Notify Party</span>
            </button>

            <button
              type="button"
              onClick={() => setShowBulkPasteModal(true)}
              className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-bold rounded border border-emerald-300 text-xs cursor-pointer transition-colors"
            >
              <ClipboardList className="w-3.5 h-3.5 text-emerald-700" />
              <span>Paste Container List (All together)</span>
            </button>

            <button
              type="button"
              onClick={handleAddContainer}
              className="inline-flex items-center gap-1 px-2.5 py-1 bg-blue-50 hover:bg-blue-100 text-[#1F4E79] font-bold rounded border border-blue-200 text-xs cursor-pointer transition-colors"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>+ Add Container</span>
            </button>

            <button
              type="button"
              onClick={handleCreateNewBill}
              className="inline-flex items-center gap-1 px-2 py-1 text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded text-xs cursor-pointer transition-colors"
              title="Start a new blank Bill of Lading"
            >
              <RotateCcw className="w-3 h-3 text-slate-500" />
              <span>New B/L</span>
            </button>
          </div>

          {/* Zoom controls */}
          <div className="flex items-center gap-2">
            <span className="text-slate-500 text-[11px]">Zoom:</span>
            <div className="inline-flex items-center bg-slate-100 rounded border border-slate-200 p-0.5">
              <button
                type="button"
                onClick={() => setZoomLevel((z) => Math.max(0.4, Number((z - 0.05).toFixed(2))))}
                className="p-1 text-slate-600 hover:text-slate-900 rounded cursor-pointer"
                title="Zoom Out"
              >
                <ZoomOut className="w-3.5 h-3.5" />
              </button>
              <span className="font-mono text-[11px] px-2 text-slate-700 min-w-[44px] text-center">
                {Math.round(zoomLevel * 100)}%
              </span>
              <button
                type="button"
                onClick={() => setZoomLevel((z) => Math.min(1.4, Number((z + 0.05).toFixed(2))))}
                className="p-1 text-slate-600 hover:text-slate-900 rounded cursor-pointer"
                title="Zoom In"
              >
                <ZoomIn className="w-3.5 h-3.5" />
              </button>
              <button
                type="button"
                onClick={() => setZoomLevel(1.0)}
                className="p-1 text-slate-500 hover:text-slate-800 rounded cursor-pointer"
                title="Reset Zoom (100%)"
              >
                <Maximize2 className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* =========================================================================
          WORKSPACE: PRINTABLE STRICT A4 BILL OF LADING SHEET
      ========================================================================== */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 flex justify-center overflow-auto">
        <div
          id="sheet-wrapper"
          style={{ transform: `scale(${zoomLevel})` }}
          className="w-full max-w-[860px] transition-transform duration-150 origin-top flex justify-center"
        >
          {/* 1-PAGE STRICT A4 BILL OF LADING SHEET */}
          <div
            id="bill-of-lading-sheet"
            className="bg-white text-slate-900 font-sans mx-auto border-2 border-[#1F4E79] shadow-xl print:shadow-none print:border-[#1F4E79] max-w-[860px] w-full select-text"
          >
            {/* 1. TOP HEADER & BOOKING REF */}
            <div className="bg-[#1F4E79] text-white p-2 flex items-center justify-between border-b border-[#1F4E79]">
              <div className="flex items-center gap-2 flex-1">
                <span className="font-extrabold text-sm sm:text-base tracking-wider text-white whitespace-nowrap">
                  BILL OF LADING —
                </span>
                <input
                  type="text"
                  value={headerTitle}
                  onChange={(e) => handleHeaderTitleChange(e.target.value)}
                  placeholder="COMPANY NAME"
                  className="font-bold text-sm sm:text-base tracking-wide text-rose-300 border-0 bg-transparent p-0 focus:outline-none flex-1"
                />
              </div>
              <div className="flex items-center gap-1 font-mono text-xs font-bold text-white shrink-0 ml-4">
                <span className="text-blue-200">B/L:</span>
                <input
                  type="text"
                  value={blNumber}
                  onChange={(e) => setBlNumber(e.target.value)}
                  placeholder="B/L Number"
                  className="font-mono font-bold text-xs text-right border-0 bg-transparent p-0 focus:outline-none text-white w-36"
                />
              </div>
            </div>

            {/* Sub-banner */}
            <div className="bg-[#D9E1F2] text-[#1F4E79] px-2.5 py-1 text-[9.5px] font-bold border-b border-[#BDD7EE] flex flex-wrap items-center justify-between gap-1">
              <div className="flex items-center gap-1">
                <span>BOOKING REF:</span>
                <input
                  type="text"
                  value={bookingRef}
                  onChange={(e) => setBookingRef(e.target.value)}
                  placeholder="BOOKING REF NO"
                  className="font-mono font-bold text-[#1F4E79] border-0 bg-transparent p-0 focus:outline-none w-56"
                />
              </div>
              <div className="flex items-center gap-1">
                <span className="text-[#1F4E79]/70">|</span>
                <span>CARRIER:</span>
                <input
                  type="text"
                  value={carrierName}
                  onChange={(e) => setCarrierName(e.target.value)}
                  placeholder="Carrier"
                  className="font-bold text-[#1F4E79] border-0 bg-transparent p-0 focus:outline-none w-52"
                />
              </div>
              <div className="flex items-center gap-1">
                <span className="text-[#1F4E79]/70">|</span>
                <span>B/L NO:</span>
                <span className="font-mono">{blNumber}</span>
              </div>
            </div>

            {/* 2. SHIPPER / REFERENCE NO ROW */}
            <div className="grid grid-cols-12 border-b border-[#1F4E79]">
              {/* Left: Shipper / Exporter */}
              <div className="col-span-8 border-r border-[#1F4E79]">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8.5px] border-b border-[#BDD7EE]">
                  SHIPPER / EXPORTER
                </div>
                <div className="p-1.5">
                  <textarea
                    rows={4}
                    value={shipperText}
                    onChange={(e) => setShipperText(e.target.value)}
                    style={{ minHeight: '80px' }}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 resize-none focus:outline-none text-[10.5px] leading-snug"
                  />
                </div>
              </div>

              {/* Right: B/L References & Freight Terms */}
              <div className="col-span-4 flex flex-col justify-between">
                <div>
                  <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8.5px] border-b border-[#BDD7EE]">
                    B/L REFERENCES & FREIGHT TERMS
                  </div>
                  <div className="p-1.5 space-y-1.5">
                    <div>
                      <span className="block font-bold text-slate-600 uppercase text-[8px]">
                        REFERENCE NO:
                      </span>
                      <input
                        type="text"
                        value={referenceNo}
                        onChange={(e) => setReferenceNo(e.target.value)}
                        className="w-full font-mono font-bold text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                      />
                    </div>
                    <div className="pt-1 border-t border-slate-200">
                      <span className="block font-bold text-slate-600 uppercase text-[8px]">
                        CARRIER:
                      </span>
                      <input
                        type="text"
                        value={carrierName}
                        onChange={(e) => setCarrierName(e.target.value)}
                        className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* 3. CONSIGNEE OR ORDER & FREIGHT TERMS */}
            <div className="grid grid-cols-12 border-b border-[#1F4E79]">
              <div className="col-span-8 border-r border-[#1F4E79]">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8.5px] border-b border-[#BDD7EE]">
                  CONSIGNEE OR ORDER
                </div>
                <div className="p-1.5">
                  <textarea
                    rows={2}
                    value={consigneeText}
                    onChange={(e) => setConsigneeText(e.target.value)}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 resize-none focus:outline-none text-[10.5px]"
                  />
                </div>
              </div>

              <div className="col-span-4 flex flex-col justify-between">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8.5px] border-b border-[#BDD7EE]">
                  FREIGHT TERMS
                </div>
                <div className="p-1.5 flex-1 flex flex-col justify-center">
                  <input
                    type="text"
                    value={freightPayableAt}
                    onChange={(e) => setFreightPayableAt(e.target.value)}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-xs"
                  />
                  <span className="text-[8.5px] font-mono text-slate-500 font-semibold mt-0.5">
                    ({movement})
                  </span>
                </div>
              </div>
            </div>

            {/* 4. FORM E NO(S) */}
            <div className="grid grid-cols-12 border-b border-[#1F4E79] bg-rose-50/20">
              <div className="col-span-12 px-2 py-1 flex items-center gap-2">
                <span className="font-bold text-[#C00000] uppercase tracking-wide text-[9.5px] whitespace-nowrap">
                  FORM E NO(s):
                </span>
                <input
                  type="text"
                  value={formENumber}
                  onChange={(e) => setFormENumber(e.target.value)}
                  className="w-full font-mono font-bold uppercase text-[#C00000] border-0 bg-transparent p-0 focus:outline-none text-[10.5px]"
                />
              </div>
            </div>

            {/* 5. DYNAMIC NOTIFY PARTIES SECTION */}
            {notifyParties.map((party, idx) => {
              if (idx % 2 !== 0) return null; // Paired with next
              const nextParty = notifyParties[idx + 1];

              return (
                <div key={party.id} className="grid grid-cols-12 border-b border-[#1F4E79]">
                  {/* First Notify Party in Row */}
                  <div className={`${nextParty ? 'col-span-6 border-r border-[#1F4E79]' : 'col-span-12'}`}>
                    <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8.5px] border-b border-[#BDD7EE] flex items-center justify-between">
                      <span>NOTIFY PARTY/ADDRESS #{idx + 1}</span>
                      <div className="flex items-center gap-1 no-print">
                        <button
                          type="button"
                          onClick={handleAddNotifyParty}
                          className="text-[8px] bg-white text-[#1F4E79] font-bold px-1 py-0.2 rounded hover:bg-blue-50 cursor-pointer"
                        >
                          + Add Party
                        </button>
                        {notifyParties.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveNotifyParty(party.id)}
                            className="text-slate-400 hover:text-rose-600 p-0.5 cursor-pointer"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        )}
                      </div>
                    </div>
                    <div className="p-1.5 space-y-1">
                      <input
                        type="text"
                        value={party.name}
                        onChange={(e) => handleUpdateNotifyParty(party.id, 'name', e.target.value)}
                        placeholder="NOTIFY PARTY NAME"
                        className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10.5px]"
                      />
                      <textarea
                        rows={3}
                        value={party.address}
                        onChange={(e) => handleUpdateNotifyParty(party.id, 'address', e.target.value)}
                        placeholder="ADDRESS, CONTACT, TEL..."
                        className="w-full font-medium uppercase text-slate-800 border-0 bg-transparent p-0 resize-none focus:outline-none text-[9.5px] leading-tight"
                      />
                    </div>
                  </div>

                  {/* Second Notify Party in Row (if exists) */}
                  {nextParty && (
                    <div className="col-span-6">
                      <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8.5px] border-b border-[#BDD7EE] flex items-center justify-between">
                        <span>NOTIFY PARTY/ADDRESS #{idx + 2}</span>
                        <div className="flex items-center gap-1 no-print">
                          <button
                            type="button"
                            onClick={handleAddNotifyParty}
                            className="text-[8px] bg-white text-[#1F4E79] font-bold px-1 py-0.2 rounded hover:bg-blue-50 cursor-pointer"
                          >
                            + Add Party
                          </button>
                          <button
                            type="button"
                            onClick={() => handleRemoveNotifyParty(nextParty.id)}
                            className="text-slate-400 hover:text-rose-600 p-0.5 cursor-pointer"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                      </div>
                      <div className="p-1.5 space-y-1">
                        <input
                          type="text"
                          value={nextParty.name}
                          onChange={(e) => handleUpdateNotifyParty(nextParty.id, 'name', e.target.value)}
                          placeholder="NOTIFY PARTY NAME"
                          className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10.5px]"
                        />
                        <textarea
                          rows={3}
                          value={nextParty.address}
                          onChange={(e) => handleUpdateNotifyParty(nextParty.id, 'address', e.target.value)}
                          placeholder="ADDRESS, CONTACT, TEL..."
                          className="w-full font-medium uppercase text-slate-800 border-0 bg-transparent p-0 resize-none focus:outline-none text-[9.5px] leading-tight"
                        />
                      </div>
                    </div>
                  )}
                </div>
              );
            })}

            {/* 6. ROUTING */}
            <div className="grid grid-cols-12 border-b border-[#1F4E79]">
              <div className="col-span-6 border-r border-[#1F4E79]">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8px] border-b border-[#BDD7EE]">
                  VESSEL AND VOY NO:
                </div>
                <div className="p-1.5">
                  <input
                    type="text"
                    value={vesselAndVoy}
                    onChange={(e) => setVesselAndVoy(e.target.value)}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
              </div>

              <div className="col-span-6">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8px] border-b border-[#BDD7EE]">
                  PLACE OF DELIVERY:
                </div>
                <div className="p-1.5">
                  <input
                    type="text"
                    value={placeOfDelivery}
                    onChange={(e) => setPlaceOfDelivery(e.target.value)}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
              </div>
            </div>

            <div className="grid grid-cols-12 border-b border-[#1F4E79]">
              <div className="col-span-6 border-r border-[#1F4E79]">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8px] border-b border-[#BDD7EE]">
                  PORT OF LOADING:
                </div>
                <div className="p-1.5">
                  <input
                    type="text"
                    value={portOfLoading}
                    onChange={(e) => setPortOfLoading(e.target.value)}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
              </div>

              <div className="col-span-6">
                <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8px] border-b border-[#BDD7EE]">
                  PORT OF DISCHARGE:
                </div>
                <div className="p-1.5">
                  <input
                    type="text"
                    value={portOfDischarge}
                    onChange={(e) => setPortOfDischarge(e.target.value)}
                    className="w-full font-bold uppercase text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
              </div>
            </div>

            {/* 7. CARGO HEADER & DESCRIPTION */}
            <div>
              <div className="grid grid-cols-12 bg-[#1F4E79] text-white border-b border-[#1F4E79]">
                <div className="col-span-3 border-r border-blue-400/40 px-2 py-1 font-bold text-[8.5px] uppercase">
                  MARKS AND NOS:
                </div>
                <div className="col-span-5 border-r border-blue-400/40 px-2 py-1 font-bold text-[8.5px] uppercase">
                  NUMBER AND KIND OF PACKAGES: DESCRIPTION OF GOODS
                </div>
                <div className="col-span-2 border-r border-blue-400/40 px-2 py-1 font-bold text-[8.5px] uppercase text-right">
                  GROSS WEIGHT
                </div>
                <div className="col-span-2 px-2 py-1 font-bold text-[8.5px] uppercase text-right">
                  MEASURMENT
                </div>
              </div>

              {/* Content Row */}
              <div className="grid grid-cols-12 border-b border-[#1F4E79]">
                <div className="col-span-3 border-r border-[#1F4E79] p-1.5 flex flex-col h-full">
                  <textarea
                    rows={14}
                    value={marksAndNumbers}
                    onChange={(e) => setMarksAndNumbers(e.target.value)}
                    style={{ minHeight: '220px' }}
                    className="w-full flex-1 font-bold uppercase text-slate-900 border-0 bg-transparent p-0 resize-none overflow-hidden focus:outline-none text-[10px] leading-snug"
                  />
                </div>

                <div className="col-span-5 border-r border-[#1F4E79] p-1.5 flex flex-col h-full">
                  <textarea
                    rows={14}
                    value={cargoDescription}
                    onChange={(e) => setCargoDescription(e.target.value)}
                    style={{ minHeight: '220px' }}
                    className="w-full flex-1 font-medium uppercase text-slate-950 border-0 bg-transparent p-0 resize-none overflow-hidden focus:outline-none text-[9px] leading-snug"
                  />
                </div>

                <div className="col-span-2 border-r border-[#1F4E79] p-1.5 flex flex-col justify-between text-right font-mono">
                  <div>
                    <input
                      type="text"
                      value={grossWeight}
                      onChange={(e) => setGrossWeight(e.target.value)}
                      className="w-full text-right font-bold text-slate-950 border-0 bg-transparent p-0 focus:outline-none text-xs"
                    />
                    <span className="text-[9px] text-[#1F4E79] font-bold block">{grossWeightUnit}</span>
                  </div>

                  <div className="my-2 pt-2 border-t border-dashed border-slate-300">
                    <span className="text-[8px] uppercase font-sans text-slate-500 font-bold block">
                      NETT WEIGHT
                    </span>
                    <input
                      type="text"
                      value={nettWeight}
                      onChange={(e) => setNettWeight(e.target.value)}
                      className="w-full text-right font-bold text-slate-950 border-0 bg-transparent p-0 focus:outline-none text-[10.5px]"
                    />
                    <span className="text-[8.5px] text-[#1F4E79] font-bold block">{nettWeightUnit}</span>
                  </div>

                  <div>
                    <input
                      type="text"
                      value={totalPackagesCount}
                      onChange={(e) => setTotalPackagesCount(e.target.value)}
                      className="w-full text-right font-bold text-slate-950 border-0 bg-transparent p-0 focus:outline-none text-[10.5px]"
                    />
                    <input
                      type="text"
                      value={totalPackagesUnit}
                      onChange={(e) => handlePackageUnitChange(e.target.value)}
                      placeholder="BAGS"
                      className="w-full text-right font-bold text-[8.5px] text-[#1F4E79] uppercase border-0 bg-transparent p-0 focus:outline-none cursor-pointer"
                    />
                  </div>
                </div>

                <div className="col-span-2 p-1.5 text-right font-mono">
                  <input
                    type="text"
                    value={measurement}
                    onChange={(e) => setMeasurement(e.target.value)}
                    className="w-full text-right text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-xs"
                  />
                </div>
              </div>

              {/* 8. EXPANDED CONTAINER MANIFEST & SPECIAL CLAUSES */}
              <div className="grid grid-cols-12 border-b border-[#1F4E79]">
                <div className="col-span-8 border-r border-[#1F4E79] flex flex-col justify-between">
                  <div>
                    <div className="bg-[#1F4E79] text-white px-2 py-1 flex items-center justify-between">
                      <div className="flex items-center gap-1.5 font-bold text-[9px] uppercase tracking-wide">
                        <span>SEAL NO: / CONTAINER NO:</span>
                        <span className="text-blue-200 font-mono font-bold">({containerCount} Cont)</span>
                      </div>

                      <div className="flex items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => setShowBulkPasteModal(true)}
                          className="no-print text-[8.5px] font-semibold text-[#1F4E79] bg-white hover:bg-blue-50 px-2 py-0.5 rounded shadow-xs flex items-center gap-1 cursor-pointer transition-all"
                        >
                          <ClipboardList className="w-3 h-3 text-[#1F4E79]" />
                          Paste List (All together)
                        </button>

                        <button
                          type="button"
                          onClick={handleAddContainer}
                          className="no-print text-[8.5px] font-semibold text-[#1F4E79] bg-white hover:bg-blue-50 px-1.5 py-0.5 rounded shadow-xs flex items-center gap-1 cursor-pointer transition-all"
                        >
                          <Plus className="w-2.5 h-2.5" />
                          Add Container
                        </button>
                      </div>
                    </div>

                    {/* Table Column Sub-headers */}
                    <div className="grid grid-cols-2 divide-x divide-[#BDD7EE] bg-[#D9E1F2] text-[#1F4E79] text-[7.5px] font-bold uppercase tracking-wider py-0.5 border-b border-[#BDD7EE]">
                      <div className="grid grid-cols-12 px-1 text-center">
                        <span className="col-span-4 text-left font-bold">Container No.</span>
                        <span className="col-span-3 text-right">Net Wt</span>
                        <span className="col-span-3 text-right">Gross Wt</span>
                        <span className="col-span-2 text-right">{totalPackagesUnit}</span>
                      </div>
                      <div className="grid grid-cols-12 px-1 text-center">
                        <span className="col-span-4 text-left font-bold">Container No.</span>
                        <span className="col-span-3 text-right">Net Wt</span>
                        <span className="col-span-3 text-right">Gross Wt</span>
                        <span className="col-span-2 text-right">{totalPackagesUnit}</span>
                      </div>
                    </div>

                    {/* 2-Column Container Rows */}
                    <div className="grid grid-cols-2 divide-x divide-[#BDD7EE] text-[8px] font-mono">
                      {/* Left Column */}
                      <div className="divide-y divide-slate-100">
                        {containers.slice(0, Math.ceil(containers.length / 2)).map((c, idx) => {
                          const hasNo = c.containerNo.trim() !== '';
                          const netDisplay = c.netWeightKg > 0 ? c.netWeightKg.toLocaleString('en-US', { minimumFractionDigits: 2 }) : (hasNo ? '26,000.00' : '');
                          const grossDisplay = c.grossWeightKg > 0 ? c.grossWeightKg.toLocaleString('en-US', { minimumFractionDigits: 2 }) : (hasNo ? '26,070.20' : '');
                          const bagsDisplay = c.packageCount > 0 ? c.packageCount : (hasNo ? '520' : '');

                          return (
                            <div key={c.id} className="grid grid-cols-12 px-1 py-0.5 items-center group hover:bg-blue-50/50">
                              <div className="col-span-4 text-left">
                                <input
                                  type="text"
                                  id={`cont-no-${idx}`}
                                  value={c.containerNo}
                                  onChange={(e) => handleContainerNoChange(idx, e.target.value)}
                                  onKeyDown={(e) => handleContainerKeyDown(idx, e)}
                                  onPaste={(e) => handleContainerPaste(idx, e)}
                                  placeholder="OSLU..."
                                  className="w-full uppercase font-mono font-bold text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[8.5px]"
                                />
                              </div>
                              <div className="col-span-3 text-right">
                                <input
                                  type="text"
                                  value={netDisplay}
                                  onChange={(e) => handleContainerWeightChange(idx, 'netWeightKg', e.target.value)}
                                  placeholder={hasNo ? '26,000.00' : ''}
                                  className="w-full text-right border-0 bg-transparent p-0 focus:outline-none text-[8px] text-slate-700"
                                />
                              </div>
                              <div className="col-span-3 text-right">
                                <input
                                  type="text"
                                  value={grossDisplay}
                                  onChange={(e) => handleContainerWeightChange(idx, 'grossWeightKg', e.target.value)}
                                  placeholder={hasNo ? '26,070.20' : ''}
                                  className="w-full text-right border-0 bg-transparent p-0 focus:outline-none text-[8px] text-slate-700"
                                />
                              </div>
                              <div className="col-span-2 text-right flex items-center justify-end gap-0.5">
                                <input
                                  type="text"
                                  value={bagsDisplay}
                                  onChange={(e) => handleContainerWeightChange(idx, 'packageCount', e.target.value)}
                                  placeholder={hasNo ? '520' : ''}
                                  className="w-full text-right border-0 bg-transparent p-0 focus:outline-none text-[8px] text-slate-800 font-semibold"
                                />
                                {containers.length > 1 && (
                                  <button
                                    type="button"
                                    onClick={() => handleRemoveContainer(c.id)}
                                    className="no-print opacity-0 group-hover:opacity-100 text-slate-400 hover:text-rose-600 px-0.5 cursor-pointer"
                                  >
                                    <Trash2 className="w-2.5 h-2.5" />
                                  </button>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>

                      {/* Right Column */}
                      <div className="divide-y divide-slate-100">
                        {containers.slice(Math.ceil(containers.length / 2)).map((c, rIdx) => {
                          const idx = Math.ceil(containers.length / 2) + rIdx;
                          const hasNo = c.containerNo.trim() !== '';
                          const netDisplay = c.netWeightKg > 0 ? c.netWeightKg.toLocaleString('en-US', { minimumFractionDigits: 2 }) : (hasNo ? '26,000.00' : '');
                          const grossDisplay = c.grossWeightKg > 0 ? c.grossWeightKg.toLocaleString('en-US', { minimumFractionDigits: 2 }) : (hasNo ? '26,070.20' : '');
                          const bagsDisplay = c.packageCount > 0 ? c.packageCount : (hasNo ? '520' : '');

                          return (
                            <div key={c.id} className="grid grid-cols-12 px-1 py-0.5 items-center group hover:bg-blue-50/50">
                              <div className="col-span-4 text-left">
                                <input
                                  type="text"
                                  id={`cont-no-${idx}`}
                                  value={c.containerNo}
                                  onChange={(e) => handleContainerNoChange(idx, e.target.value)}
                                  onKeyDown={(e) => handleContainerKeyDown(idx, e)}
                                  onPaste={(e) => handleContainerPaste(idx, e)}
                                  placeholder="OSLU..."
                                  className="w-full uppercase font-mono font-bold text-slate-900 border-0 bg-transparent p-0 focus:outline-none text-[8.5px]"
                                />
                              </div>
                              <div className="col-span-3 text-right">
                                <input
                                  type="text"
                                  value={netDisplay}
                                  onChange={(e) => handleContainerWeightChange(idx, 'netWeightKg', e.target.value)}
                                  placeholder={hasNo ? '26,000.00' : ''}
                                  className="w-full text-right border-0 bg-transparent p-0 focus:outline-none text-[8px] text-slate-700"
                                />
                              </div>
                              <div className="col-span-3 text-right">
                                <input
                                  type="text"
                                  value={grossDisplay}
                                  onChange={(e) => handleContainerWeightChange(idx, 'grossWeightKg', e.target.value)}
                                  placeholder={hasNo ? '26,070.20' : ''}
                                  className="w-full text-right border-0 bg-transparent p-0 focus:outline-none text-[8px] text-slate-700"
                                />
                              </div>
                              <div className="col-span-2 text-right flex items-center justify-end gap-0.5">
                                <input
                                  type="text"
                                  value={bagsDisplay}
                                  onChange={(e) => handleContainerWeightChange(idx, 'packageCount', e.target.value)}
                                  placeholder={hasNo ? '520' : ''}
                                  className="w-full text-right border-0 bg-transparent p-0 focus:outline-none text-[8px] text-slate-800 font-semibold"
                                />
                                {containers.length > 1 && (
                                  <button
                                    type="button"
                                    onClick={() => handleRemoveContainer(c.id)}
                                    className="no-print opacity-0 group-hover:opacity-100 text-slate-400 hover:text-rose-600 px-0.5 cursor-pointer"
                                  >
                                    <Trash2 className="w-2.5 h-2.5" />
                                  </button>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>

                  {/* Manifest Totals Bar */}
                  <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 text-[8.5px] font-bold border-t border-[#BDD7EE] flex items-center justify-between">
                    <span>TOTALS: {containerCount} CONTAINER(S)</span>
                    <div className="flex items-center gap-2 font-mono">
                      <span>NET: {totalNetKg.toLocaleString('en-US', { minimumFractionDigits: 2 })} KG</span>
                      <span className="text-[#1F4E79]/50">|</span>
                      <span>GROSS: {totalGrossKg.toLocaleString('en-US', { minimumFractionDigits: 2 })} KG</span>
                      <span className="text-[#1F4E79]/50">|</span>
                      <span>{totalBags.toLocaleString()} {totalPackagesUnit.toUpperCase()}</span>
                    </div>
                  </div>
                </div>

                {/* Right Column: Special Instructions / Clauses */}
                <div className="col-span-4 p-2 flex flex-col justify-between bg-white relative">
                  <div>
                    <div className="bg-[#D9E1F2] text-[#1F4E79] px-2 py-0.5 font-bold uppercase text-[8px] border-b border-[#BDD7EE] mb-1.5 -mx-2 -mt-2 flex items-center justify-between">
                      <span>SPECIAL INSTRUCTIONS / CLAUSES</span>
                      <button
                        type="button"
                        onClick={() => setShowClausesPopover(!showClausesPopover)}
                        className="no-print p-0.5 hover:bg-[#BDD7EE] rounded text-[#1F4E79] cursor-pointer transition-colors"
                        title="Configure clauses and presets"
                      >
                        <MoreVertical className="w-3 h-3" />
                      </button>
                    </div>

                    {/* Clauses Popover */}
                    {showClausesPopover && (
                      <div className="no-print absolute top-7 right-1 z-50 w-72 bg-white rounded-lg shadow-2xl border border-slate-300 p-3 text-left">
                        <div className="flex items-center justify-between pb-2 mb-2 border-b border-slate-200">
                          <span className="font-bold text-xs text-[#1F4E79] uppercase tracking-wide">
                            Select Clauses / Presets
                          </span>
                          <button
                            type="button"
                            onClick={() => setShowClausesPopover(false)}
                            className="text-slate-400 hover:text-slate-600 p-0.5 cursor-pointer"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        </div>

                        <div className="space-y-1 mb-3">
                          <span className="text-[10px] font-bold text-slate-500 uppercase block">1-Click Presets:</span>
                          <button
                            type="button"
                            onClick={() => handleApplyPreset('sugar')}
                            className="w-full text-left px-2 py-1.5 rounded text-xs font-semibold bg-slate-50 hover:bg-blue-50 hover:text-[#1F4E79] border border-slate-200 flex items-center justify-between cursor-pointer transition-colors"
                          >
                            <span>Standard Sugar / Dry Cargo</span>
                            <span className="text-[9px] text-slate-400 font-mono">(5)</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => handleApplyPreset('reefer')}
                            className="w-full text-left px-2 py-1.5 rounded text-xs font-semibold bg-blue-50/70 text-[#1F4E79] border border-blue-300 flex items-center justify-between cursor-pointer transition-colors"
                          >
                            <span>Reefer Cargo (-18°C)</span>
                            <span className="text-[9px] text-blue-600 font-mono font-bold">(7)</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => handleApplyPreset('general')}
                            className="w-full text-left px-2 py-1.5 rounded text-xs font-semibold bg-slate-50 hover:bg-blue-50 hover:text-[#1F4E79] border border-slate-200 flex items-center justify-between cursor-pointer transition-colors"
                          >
                            <span>General Containerized Cargo</span>
                            <span className="text-[9px] text-slate-400 font-mono">(6)</span>
                          </button>
                        </div>

                        <div className="mb-3">
                          <span className="text-[10px] font-bold text-slate-500 uppercase block mb-1">Toggle Clauses:</span>
                          <div className="max-h-36 overflow-y-auto space-y-1 text-[11px] pr-1">
                            {ALL_AVAILABLE_CLAUSES.map((cl) => {
                              const active = specialClauses.includes(cl);
                              return (
                                <label
                                  key={cl}
                                  className={`flex items-start gap-1.5 p-1 rounded cursor-pointer text-left select-none ${
                                    active ? 'bg-blue-50 text-[#1F4E79] font-bold' : 'text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  <input
                                    type="checkbox"
                                    checked={active}
                                    onChange={() => handleToggleClause(cl)}
                                    className="mt-0.5 rounded text-[#1F4E79]"
                                  />
                                  <span className="flex-1">{cl}</span>
                                </label>
                              );
                            })}
                          </div>
                        </div>

                        <div className="pt-2 border-t border-slate-200 flex gap-1">
                          <input
                            type="text"
                            value={customClauseInput}
                            onChange={(e) => setCustomClauseInput(e.target.value)}
                            placeholder="Add custom clause..."
                            className="flex-1 text-xs border border-slate-300 rounded px-2 py-1 uppercase"
                          />
                          <button
                            type="button"
                            onClick={handleAddCustomClause}
                            className="px-2 py-1 bg-[#1F4E79] text-white text-xs font-bold rounded cursor-pointer"
                          >
                            Add
                          </button>
                        </div>
                      </div>
                    )}

                    {/* Display of Active Clauses */}
                    <div className="space-y-1 font-bold text-slate-900 uppercase tracking-tight text-[9px]">
                      {specialClauses.map((clause, idx) => (
                        <div key={idx} className="leading-tight">
                          {clause}
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="mt-3 pt-1 border-t border-[#BDD7EE]">
                    <textarea
                      rows={2}
                      value={containerSummaryLine}
                      onChange={(e) => setContainerSummaryLine(e.target.value)}
                      style={{ minHeight: '34px' }}
                      className="w-full font-bold uppercase text-[#1F4E79] border-0 bg-transparent p-0 resize-none overflow-hidden focus:outline-none text-[8.5px] leading-snug"
                    />
                  </div>
                </div>
              </div>

              <div className="bg-[#D9E1F2] text-[#1F4E79] py-0.5 px-2 border-b border-[#1F4E79] text-center text-[8.5px] font-bold tracking-wider">
                ABOVE PARTICULARS AS DECLARED BY SHIPPER
              </div>
            </div>

            {/* 9. BOTTOM FOOTER GRID */}
            <div className="grid grid-cols-12 divide-x divide-[#1F4E79] text-slate-900 border-t border-[#1F4E79]">
              <div className="col-span-5 divide-y divide-[#1F4E79]">
                <div className="p-1.5">
                  <span className="block font-bold text-[#1F4E79] uppercase tracking-tight text-[8px]">
                    TOTAL NO. OF CONTAINERS/PACKAGES
                  </span>
                  <input
                    type="text"
                    value={totalContainersPackages}
                    onChange={(e) => setTotalContainersPackages(e.target.value)}
                    placeholder="1X20' FCL"
                    className="w-full font-mono font-bold uppercase border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
                <div className="p-1.5">
                  <span className="block font-bold text-[#1F4E79] uppercase tracking-tight text-[8px]">
                    MOVEMENT
                  </span>
                  <input
                    type="text"
                    value={movement}
                    onChange={(e) => setMovement(e.target.value)}
                    placeholder="FCL/FCL"
                    className="w-full font-mono font-bold uppercase border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
              </div>

              <div className="col-span-4 divide-y divide-[#1F4E79]">
                <div className="p-1.5">
                  <span className="block font-bold text-[#1F4E79] uppercase tracking-tight text-[8px]">
                    FREIGHT PAYABLE AT
                  </span>
                  <input
                    type="text"
                    value={freightPayableAt}
                    onChange={(e) => setFreightPayableAt(e.target.value)}
                    placeholder="FREIGHT PREPAID"
                    className="w-full font-mono font-bold uppercase border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
                <div className="p-1.5">
                  <span className="block font-bold text-[#1F4E79] uppercase tracking-tight text-[8px]">
                    NUMBER OF ORIGINAL B/L
                  </span>
                  <input
                    type="text"
                    value={numberOfOriginals}
                    onChange={(e) => setNumberOfOriginals(e.target.value)}
                    placeholder="(3) THREE"
                    className="w-full font-mono font-bold uppercase border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
              </div>

              <div className="col-span-3 p-1.5 flex flex-col justify-between">
                <div>
                  <span className="block font-bold text-[#1F4E79] uppercase tracking-tight text-[8px]">
                    PLACE AND DATE OF ISSUE
                  </span>
                  <input
                    type="text"
                    value={placeAndDateOfIssue}
                    onChange={(e) => setPlaceAndDateOfIssue(e.target.value)}
                    placeholder="KARACHI ......"
                    className="w-full font-mono font-bold uppercase border-0 bg-transparent p-0 focus:outline-none text-[10px]"
                  />
                </div>
                <div className="pt-3">
                  <span className="block font-bold text-[#1F4E79] uppercase text-[7.5px] tracking-tight">
                    AS CARRIER / AUTHORIZED AGENT
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* =========================================================================
          GOOGLE CLOUD SAVED BILLS MODAL
      ========================================================================== */}
      {showSavedModal && (
        <div className="no-print fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-2xl w-full max-h-[85vh] flex flex-col text-slate-900 overflow-hidden">
            <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-lg bg-blue-100 text-[#1F4E79] flex items-center justify-center">
                  <Database className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                    Google Cloud Saved Bills of Lading
                    <span className="text-[10px] bg-blue-100 text-blue-800 font-mono px-2 py-0.5 rounded-full font-semibold">
                      {savedBills.length} Saved in Account
                    </span>
                  </h3>
                  <p className="text-[11px] text-slate-500">
                    {currentUser
                      ? `Linked to Google account: ${currentUser.email}`
                      : 'Sign in to sync your saved documents in Google Cloud'}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowSavedModal(false)}
                className="text-slate-400 hover:text-slate-600 p-1.5 rounded-lg hover:bg-slate-200/60 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Search filter & Actions */}
            <div className="p-3 border-b border-slate-200 bg-white flex items-center gap-2">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={savedSearchQuery}
                  onChange={(e) => setSavedSearchQuery(e.target.value)}
                  placeholder="Search by B/L Number, Company Name, Carrier..."
                  className="w-full text-xs pl-9 pr-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>
              <button
                type="button"
                onClick={handleCreateNewBill}
                className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold rounded-lg flex items-center gap-1 cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>New Blank</span>
              </button>
            </div>

            {/* List */}
            <div className="flex-1 overflow-y-auto p-4 space-y-2.5 bg-slate-50/50 min-h-[220px]">
              {!currentUser ? (
                <div className="text-center py-10 text-slate-500 text-xs">
                  <Cloud className="w-10 h-10 mx-auto mb-2 text-blue-400 opacity-60" />
                  <p className="font-semibold text-slate-700">Please Sign In with Google</p>
                  <p className="text-[11px] text-slate-400 mt-1 max-w-sm mx-auto">
                    Sign in with your Google account so your Bills of Lading and profile data persist permanently in Google Cloud Storage.
                  </p>
                  <button
                    type="button"
                    onClick={handleGoogleLogin}
                    className="mt-3 px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-semibold rounded-lg text-xs cursor-pointer shadow inline-flex items-center gap-1.5"
                  >
                    <LogIn className="w-4 h-4" />
                    Sign In with Google
                  </button>
                </div>
              ) : filteredSavedBills.length === 0 ? (
                <div className="text-center py-10 text-slate-400 text-xs">
                  <Database className="w-8 h-8 mx-auto mb-2 opacity-50" />
                  <p>No Bills of Lading saved in Google Cloud yet.</p>
                  <p className="text-[10px] text-slate-400 mt-1">
                    Click &ldquo;Save to Cloud&rdquo; on the top bar to save your current document.
                  </p>
                </div>
              ) : (
                filteredSavedBills.map((item) => (
                  <div
                    key={item.id}
                    onClick={() => handleLoadSavedBill(item)}
                    className="bg-white p-3 rounded-lg border border-slate-200 hover:border-blue-400 shadow-xs hover:shadow transition-all flex items-center justify-between gap-3 cursor-pointer group"
                  >
                    <div>
                      <div className="flex items-center gap-2">
                        <span className="font-mono font-bold text-xs text-[#1F4E79]">{item.blNumber}</span>
                        <span className="text-[10px] bg-blue-50 text-blue-700 px-1.5 py-0.5 rounded font-semibold">
                          {item.containers ? item.containers.length : 0} Containers
                        </span>
                        <span className="text-[10px] text-slate-400">
                          {item.savedAt ? `Saved: ${item.savedAt}` : ''}
                        </span>
                      </div>
                      <p className="text-xs font-semibold text-slate-800 mt-0.5">{item.headerTitle || 'N/A'}</p>
                      <p className="text-[10px] text-slate-500">
                        {item.carrierName} &bull; POD: {item.portOfDischarge}
                      </p>
                    </div>

                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => handleLoadSavedBill(item)}
                        className="px-3 py-1.5 bg-blue-50 hover:bg-blue-100 text-[#1F4E79] text-xs font-bold rounded cursor-pointer"
                      >
                        Load
                      </button>
                      <button
                        type="button"
                        onClick={(e) => handleDeleteSavedBill(e, item.id, item.blNumber)}
                        className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded cursor-pointer"
                        title="Delete from Google Cloud"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* Footer */}
            <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
              <div className="flex items-center gap-1.5 text-emerald-700 font-semibold text-[11px]">
                <CheckCircle className="w-3.5 h-3.5 text-emerald-600" />
                <span>Google Cloud Firestore Persistence Active</span>
              </div>
              <button
                type="button"
                onClick={() => setShowSavedModal(false)}
                className="px-3.5 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold rounded-md transition-colors cursor-pointer"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* =========================================================================
          BULK PASTE CONTAINERS MODAL
      ========================================================================== */}
      {showBulkPasteModal && (
        <div className="no-print fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-lg w-full p-4 sm:p-5 text-left">
            <div className="flex items-center justify-between pb-3 border-b border-slate-200">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 rounded-lg bg-blue-100 text-[#1F4E79] flex items-center justify-center font-bold">
                  <ClipboardList className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-sm font-bold text-slate-900">Paste Containers (All together)</h3>
                  <p className="text-[11px] text-slate-500">
                    Paste full container list from Excel, WhatsApp, or Notepad
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowBulkPasteModal(false)}
                className="text-slate-400 hover:text-slate-600 p-1 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="py-3">
              <label className="block text-xs font-bold text-slate-700 mb-1">
                Paste Container Numbers (1 per line, comma, or tab separated):
              </label>
              <textarea
                rows={9}
                value={bulkPasteText}
                onChange={(e) => setBulkPasteText(e.target.value)}
                placeholder={'BISU2834268\nCAIU2369026\nCCLU3517766\nCCLU3802207\nCLHU3363830\nEGHU3048530\nGESU3716288\nGLDU3067966\n...'}
                className="w-full font-mono text-xs uppercase p-2.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
              />
              <div className="flex items-center justify-between mt-1 text-[11px]">
                <span className="font-semibold text-blue-700">
                  Detected:{' '}
                  {
                    bulkPasteText
                      .split(/[\r\n\t,]+/)
                      .map((s) => s.trim())
                      .filter((s) => s.length > 0).length
                  }{' '}
                  containers
                </span>
                <label className="flex items-center gap-1.5 cursor-pointer text-slate-600">
                  <input
                    type="checkbox"
                    checked={bulkReplaceExisting}
                    onChange={(e) => setBulkReplaceExisting(e.target.checked)}
                    className="rounded text-[#1F4E79]"
                  />
                  <span>Replace existing containers</span>
                </label>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 pt-3 border-t border-slate-200">
              <button
                type="button"
                onClick={() => setShowBulkPasteModal(false)}
                className="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleApplyBulkContainers}
                className="px-4 py-1.5 text-xs font-bold text-white bg-[#1F4E79] hover:bg-[#163a5c] rounded-lg shadow cursor-pointer flex items-center gap-1.5"
              >
                <Check className="w-3.5 h-3.5" />
                Apply & Auto-Fill Weights
              </button>
            </div>
          </div>
        </div>
      )}

      {/* =========================================================================
          SIGN IN PROMPT MODAL
      ========================================================================== */}
      {showSignInPromptModal && (
        <div className="no-print fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-md w-full p-6 text-center">
            <div className="w-12 h-12 rounded-full bg-blue-100 text-[#1F4E79] mx-auto flex items-center justify-center mb-3">
              <Cloud className="w-6 h-6 text-blue-700" />
            </div>
            <h3 className="text-base font-bold text-slate-900 mb-1">Save Permanently to Google Cloud</h3>
            <p className="text-xs text-slate-600 mb-4 leading-relaxed">
              Apne Bills of Lading ko permanent Google Cloud Firestore me save karne ke liye apne Google account se sign in karein. Jab bhi aap logout karke dobara login karenge, aapka sara data wapis aa jayega.
            </p>

            <button
              type="button"
              onClick={handleGoogleLogin}
              disabled={authLoading}
              className="w-full py-2.5 px-4 bg-white hover:bg-slate-50 text-slate-800 font-bold text-xs rounded-xl border border-slate-300 shadow flex items-center justify-center gap-2 cursor-pointer transition-all mb-2"
            >
              <svg className="w-4 h-4 shrink-0" viewBox="0 0 24 24">
                <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/>
                <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/>
                <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z"/>
                <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z"/>
              </svg>
              <span>{authLoading ? 'Connecting...' : 'Continue with Google'}</span>
            </button>

            <button
              type="button"
              onClick={() => {
                setShowSignInPromptModal(false);
                setPendingCloudSave(false);
              }}
              className="text-xs text-slate-500 hover:text-slate-800 font-medium py-1 cursor-pointer"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {/* =========================================================================
          TOAST NOTIFICATION
      ========================================================================== */}
      {toastMessage && (
        <div
          className={`no-print fixed bottom-6 right-6 z-50 px-4 py-3 rounded-xl shadow-xl flex items-center gap-2.5 text-xs font-semibold border transition-all ${
            toastMessage.type === 'error'
              ? 'bg-rose-900 text-white border-rose-700'
              : toastMessage.type === 'info'
              ? 'bg-slate-900 text-white border-slate-700'
              : 'bg-emerald-900 text-white border-emerald-700'
          }`}
        >
          {toastMessage.type === 'error' ? (
            <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
          ) : (
            <CheckCircle className="w-4 h-4 text-emerald-400 shrink-0" />
          )}
          <span>{toastMessage.text}</span>
        </div>
      )}
    </div>
  );
}
