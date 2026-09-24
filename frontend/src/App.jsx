import React, { useState, useRef, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import './index.css';
import logo from './assets/logo.svg';

// Función para reproducir sonidos sin necesidad de archivos de audio
const playSound = (type = 'error') => {
  try {
    const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    const oscillator = audioCtx.createOscillator();
    const gainNode = audioCtx.createGain();
    
    oscillator.connect(gainNode);
    gainNode.connect(audioCtx.destination);
    
    if (type === 'error') {
      // Sonido de error grave (Buzz)
      oscillator.type = 'sawtooth';
      oscillator.frequency.setValueAtTime(150, audioCtx.currentTime);
      gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
      oscillator.start();
      oscillator.stop(audioCtx.currentTime + 0.3);
    } else {
      // Sonido de éxito agudo (Ping)
      oscillator.type = 'sine';
      oscillator.frequency.setValueAtTime(800, audioCtx.currentTime);
      gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
      oscillator.start();
      oscillator.stop(audioCtx.currentTime + 0.1);
    }
  } catch (e) {
    console.log("Audio API no soportada", e);
  }
};

function App() {
  const [data, setData] = useState([]);
  const [scanInput, setScanInput] = useState('');
  const [lastOrderScanned, setLastOrderScanned] = useState('');
  const [fileName, setFileName] = useState('');
  const [autoFocusEnabled, setAutoFocusEnabled] = useState(true);
  
  // Sistemas de Notificación y Auditoría
  const [toast, setToast] = useState(null);
  const [unknownGuia, setUnknownGuia] = useState(null);
  const [scanHistory, setScanHistory] = useState([]); // Historial inmutable
  const [showHistory, setShowHistory] = useState(false); // Modal de historial
  const [isDragging, setIsDragging] = useState(false); // Drag & drop state

  const scanInputRef = useRef(null);
  const toastTimerRef = useRef(null);

  // Muestra un cartel lateral no bloqueante
  const showToast = (message, type = 'success') => {
    playSound(type);
    setToast({ message, type });
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => setToast(null), 2500);
  };

  // Mantener el foco
  useEffect(() => {
    if (autoFocusEnabled && data.length > 0 && scanInputRef.current && !unknownGuia && !showHistory) {
      scanInputRef.current.focus();
    }
  }, [data, autoFocusEnabled, unknownGuia, showHistory]);

  const handleWorkspaceClick = (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON' || e.target.tagName === 'LABEL') return;
    if (autoFocusEnabled && scanInputRef.current && !unknownGuia && !showHistory) {
      scanInputRef.current.focus();
    }
  };

  const processJsonData = (jsonData, name) => {
    if (!jsonData || jsonData.length === 0) return;
    
    const match = name.match(/_(\d{5})/);
    if (match) {
      setFileName(match[1]);
    } else {
      setFileName(name.split('.')[0].substring(0, 8));
    }

    const parsedData = jsonData.map((row, index) => {
      const getVal = (possibleKeys) => {
        const key = Object.keys(row).find(k => possibleKeys.includes(k.trim().toUpperCase()));
        return key ? row[key] : '';
      };

      const guia = getVal(['GUIA']);
      const order = getVal(['O']);
      const bultosRealesRaw = getVal(['BULTOS REALES', 'BULTOS']);
      const bultosReales = parseInt(bultosRealesRaw, 10) || 1;
      
      return {
        id: index.toString(),
        guia: String(guia).trim(),
        order: String(order).trim(),
        control: 'PENDIENTE',
        estadoControl: '',
        numeroOrden: '',
        scaner: 0,
        controlBultos: 0,
        bultosReales: bultosReales,
        bultos: bultosReales,
        servicio: getVal(['SERVICIO']),
        cliente: getVal(['CLIENTE']),
        destinatario: getVal(['DESTINATARIO'])
      };
    });
    
    setData(parsedData);
    setLastOrderScanned('0');
  };

  const processFile = (file) => {
    if (!file) return;
    const extension = file.name.split('.').pop().toLowerCase();

    if (extension === 'csv') {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        encoding: "UTF-8",
        complete: (results) => processJsonData(results.data, file.name),
        error: () => showToast('Error al leer el archivo CSV', 'error')
      });
    } else if (extension === 'xls' || extension === 'xlsx') {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames.find(n => n.toUpperCase() === 'PLANILLA') || workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          processJsonData(jsonData, file.name);
        } catch(err) {
          showToast('Error al leer el Excel', 'error');
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      showToast('Formato no soportado. Use .csv, .xls o .xlsx', 'error');
    }
  };

  const handleFileUpload = (event) => {
    processFile(event.target.files[0]);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleScan = (e) => {
    if (e.key === 'Enter') {
      const codeToFind = scanInput.trim();
      if (!codeToFind) return;

      setScanInput(''); // Limpiar inmediatamente

      const itemIndex = data.findIndex(item => item.guia.toLowerCase() === codeToFind.toLowerCase());

      if (itemIndex !== -1) {
        // ENCONTRADO
        const newData = [...data];
        const item = newData[itemIndex];
        
        item.scaner = 0;
        item.controlBultos += 1;
        
        let eventType = 'INCOMPLETO (Faltan bultos)';

        if (item.controlBultos === item.bultosReales) {
          item.control = 'OK';
          eventType = 'CONTROL OK';
          showToast(`Guía Completada: ${item.guia}`, 'success');
        } else if (item.controlBultos > item.bultosReales) {
          item.control = 'EXCESO';
          eventType = 'EXCESO';
          showToast(`Exceso en GUIA: ${item.guia}`, 'error');
        } else {
          item.control = 'PENDIENTE';
          showToast(`Bulto ${item.controlBultos} de ${item.bultosReales} (${item.guia})`, 'success');
        }

        // Auditoría inmutable
        setScanHistory(prev => [{ 
          time: new Date().toLocaleTimeString(), 
          guia: item.guia, 
          motivo: eventType,
          servicio: item.servicio,
          controlBultos: item.controlBultos,
          bultosReales: item.bultosReales
        }, ...prev]);

        setLastOrderScanned(item.order);
        setData(newData);
      } else {
        // NO ENCONTRADO - Mostrar modal bloqueante
        playSound('error');
        setUnknownGuia(codeToFind);
      }
    }
  };

  // Manejo de las opciones del modal bloqueante
  const handleUnknownAction = (action) => {
    if (action === 'agregar') {
      const newRow = {
        id: Date.now().toString(),
        guia: unknownGuia,
        order: 'NUEVO',
        control: 'OK',
        estadoControl: 'FUERA DE PLANILLA',
        numeroOrden: '',
        scaner: 1, // Ya escaneó el primer bulto
        controlBultos: 1,
        bultosReales: 1,
        bultos: 1,
        servicio: 'Extra',
        cliente: 'Sin cliente',
        destinatario: 'Sin destino'
      };
      setData([...data, newRow]);
      setLastOrderScanned('NUEVO');
      setScanHistory(prev => [{ 
        time: new Date().toLocaleTimeString(), 
        guia: unknownGuia, 
        motivo: 'AGREGADO (Fuera de Planilla)',
        servicio: 'Desconocido',
        controlBultos: 1,
        bultosReales: 1
      }, ...prev]);
      showToast('Guía agregada a la lista', 'success');
    } else {
      // Acción 'separar' = simplemente se ignora y no se suma a la tabla
      setScanHistory(prev => [{ 
        time: new Date().toLocaleTimeString(), 
        guia: unknownGuia, 
        motivo: 'SEPARADO (Ignorado)',
        servicio: '-',
        controlBultos: 1,
        bultosReales: 1
      }, ...prev]);
      showToast('Paquete separado y omitido', 'error');
    }
    
    setUnknownGuia(null);
  };

  const clearData = () => {
    if (window.confirm('¿Desea limpiar el sistema y cargar una nueva planilla?')) {
      setData([]);
      setLastOrderScanned('');
      setFileName('');
      setScanHistory([]);
    }
  };

  if (data.length === 0) {
    return (
      <div className="app-container">
        <div 
          className={`upload-zone ${isDragging ? 'dragging' : ''}`} 
          onClick={() => document.getElementById('file-upload').click()}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <h2 style={{ color: 'var(--header-blue)', fontSize: '2rem', marginBottom: '10px' }}>NUEVA PLANILLA</h2>
          <p style={{ color: '#555', fontSize: '1.2rem', pointerEvents: 'none' }}>
            {isDragging ? 'Suelta el archivo aquí...' : 'Haga clic o arrastre el archivo .XLS, .XLSX o .CSV aquí'}
          </p>
          <input type="file" id="file-upload" style={{ display: 'none' }} accept=".csv, .xls, .xlsx" onChange={handleFileUpload} />
        </div>
      </div>
    );
  }

  const guiasPendientes = data.filter(d => d.control !== 'OK' && d.control !== 'EXCESO').length;
  const estadoEscaneo = data.length - guiasPendientes;

  return (
    <div className="app-container" onClick={handleWorkspaceClick}>
      
      {/* Modal Bloqueante para Guía No Encontrada */}
      {unknownGuia && (
        <div className="modal-overlay">
          <div className="modal-content">
            <h2 style={{ color: '#ef4444', marginBottom: '15px', fontSize: '2.5rem' }}>¡GUÍA NO CORRESPONDE!</h2>
            <p style={{ fontSize: '1.3rem', marginBottom: '30px' }}>
              La guía <strong style={{fontSize: '1.8rem'}}>{unknownGuia}</strong> no existe en esta planilla.
            </p>
            <div style={{ display: 'flex', gap: '20px', justifyContent: 'center' }}>
              <button className="btn" onClick={() => handleUnknownAction('agregar')} style={{ background: '#00b0f0', color: 'white', padding: '15px 30px', fontSize: '1.2rem' }}>
                AGREGARLA (FUERA DE PLANILLA)
              </button>
              <button className="btn" onClick={() => handleUnknownAction('separar')} style={{ background: '#ef4444', color: 'white', padding: '15px 30px', fontSize: '1.2rem' }}>
                SEPARAR PAQUETE (IGNORAR)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de Historial Inmutable */}
      {showHistory && (
        <div className="modal-overlay" onClick={() => setShowHistory(false)}>
          <div className="modal-content history-modal" onClick={e => e.stopPropagation()}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <h2 style={{ color: 'var(--header-blue)' }}>Historial de Escaneo (Solo Lectura)</h2>
              <button onClick={() => setShowHistory(false)} style={{ background: 'none', border: 'none', fontSize: '1.5rem', cursor: 'pointer', color: 'red' }}>✖</button>
            </div>
            
            <p style={{textAlign: 'left', marginBottom: '15px', fontSize: '0.9rem', color: '#666'}}>
              Este registro es inmutable. Registra cada escaneo y error para auditoría de los operarios.
            </p>

            <div className="history-table-container">
              <table style={{ width: '100%', textAlign: 'center' }}>
                <thead>
                  <tr>
                    <th style={{textAlign: 'left'}}>Hora</th>
                    <th style={{textAlign: 'left'}}>Guía</th>
                    <th style={{textAlign: 'left'}}>Motivo</th>
                    <th>C. Bultos</th>
                    <th>B. Reales</th>
                    <th style={{textAlign: 'left'}}>Servicio</th>
                  </tr>
                </thead>
                <tbody>
                  {scanHistory.length === 0 ? (
                    <tr><td colSpan="6" style={{ textAlign: 'center', padding: '20px' }}>No hay escaneos registrados aún.</td></tr>
                  ) : (
                    scanHistory.map((log, idx) => (
                      <tr key={idx} className={
                        log.motivo.includes('EXCESO') ? 'status-error' : 
                        log.motivo.includes('SEPARADO') ? 'status-error' : 
                        log.motivo.includes('Faltan') ? 'status-error' : 
                        log.motivo.includes('OK') ? 'status-ok' : ''
                      }>
                        <td style={{textAlign: 'left'}}>{log.time}</td>
                        <td style={{ fontWeight: 'bold', textAlign: 'left' }}>{log.guia}</td>
                        <td style={{ fontWeight: 'bold', textAlign: 'left' }}>{log.motivo}</td>
                        <td style={{ fontWeight: 'bold', fontSize: '1.1rem' }}>{log.controlBultos}</td>
                        <td style={{ fontWeight: 'bold', fontSize: '1.1rem' }}>{log.bultosReales}</td>
                        <td style={{textAlign: 'left'}}>{log.servicio}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Toast no bloqueante */}
      {toast && (
        <div className={`toast-notification ${toast.type}`}>
          {toast.message}
        </div>
      )}

      <div className="dashboard-top">
        <div className="logo-area">
          <img src={logo} alt="Fixy Logística" style={{ width: '100%', maxHeight: '60px', objectFit: 'contain' }} />
        </div>
        
        <div className="stats-container">
          <div className="stat-box">
            <div className="stat-header">ORDEN</div>
            <div className="stat-value">{lastOrderScanned || '0'}</div>
          </div>
          <div className="stat-box">
            <div className="stat-header">ESTADO DE ESCANEO</div>
            <div className="stat-value">{estadoEscaneo}</div>
          </div>
          <div className="stat-box">
            <div className="stat-header">GUIAS PENDIENTES</div>
            <div className="stat-value">{guiasPendientes}</div>
          </div>
          <div className="stat-box">
            <div className="stat-header">PLANILLA</div>
            <div className="stat-value">{fileName}</div>
          </div>
          <div className="stat-box">
            <div className="stat-header">TRANSPORTISTA</div>
            <div className="stat-value"></div>
          </div>
        </div>
      </div>

      <div className="action-bar">
        <button className="btn">Actualizar Transportistas</button>
        <button className="btn">Nueva Planilla</button>
        <button className="btn" onClick={clearData}>Limpieza</button>
        <button className="btn">Guardar Archivo</button>
        
        <button className="btn" onClick={() => setShowHistory(true)} style={{ background: '#f59e0b', color: 'white', border: 'none', marginLeft: '10px' }}>
          Historial de Control
        </button>
        
        <div style={{ flex: 1 }}></div>

        <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontWeight: 600, fontSize: '0.9rem', color: 'var(--text-main)' }}>
          <input 
            type="checkbox" 
            checked={autoFocusEnabled}
            onChange={(e) => {
              setAutoFocusEnabled(e.target.checked);
              if (e.target.checked && scanInputRef.current) scanInputRef.current.focus();
            }}
            style={{ width: '18px', height: '18px', cursor: 'pointer' }}
          />
          Forzar foco escáner
        </label>

        <input
          ref={scanInputRef}
          type="text"
          className="scan-input"
          placeholder="Escanee la GUIA aquí..."
          value={scanInput}
          onChange={(e) => setScanInput(e.target.value)}
          onKeyDown={handleScan}
          disabled={unknownGuia !== null}
        />
      </div>

      <div className="table-container">
        <table>
          <thead>
            <tr>
              <th>GUIA</th>
              <th>O</th>
              <th>CONTROL</th>
              <th>ESTADO DE CONTROL</th>
              <th>NUMERO DE ORDEN</th>
              <th>CONTROL BULTOS</th>
              <th>BULTOS REALES</th>
              <th>BULTOS</th>
              <th>SERVICIO</th>
              <th>CLIENTE</th>
              <th>DESTINATARIO</th>
            </tr>
          </thead>
          <tbody>
            {data.map((item) => (
              <tr 
                key={item.id} 
                className={item.control === 'OK' ? 'status-ok' : item.control === 'EXCESO' ? 'status-error' : ''}
              >
                <td>{item.guia}</td>
                <td>{item.order}</td>
                <td>{item.control}</td>
                <td>{item.estadoControl}</td>
                <td>{item.numeroOrden}</td>
                <td>{item.controlBultos}</td>
                <td>{item.bultosReales}</td>
                <td>{item.bultos}</td>
                <td>{item.servicio}</td>
                <td>{item.cliente}</td>
                <td>{item.destinatario}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
    </div>
  );
}

export default App;
