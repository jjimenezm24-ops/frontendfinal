import React, { useState, useRef } from 'react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import {
  getArticulosFiltroNombre,
  getReporteArticuloCategoria,
  getReporteClienteFactura
} from '../services/reportesService';
import { Articulo, ArticuloCategoriaDTO, ClienteFacturaDTO } from '../types/reportes';

export const Reportes: React.FC = () => {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [filtroNombre, setFiltroNombre] = useState<string>('');
  const [articulosFiltrados, setArticulosFiltrados] = useState<Articulo[]>([]);

  // Configuración base de PDF (Hoja A4 Horizontal, título, fecha y encabezado)
  const crearPDFBase = (tituloReporte: string) => {
    const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const fechaActual = new Date().toLocaleDateString('es-GT', {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });

    // Encabezado corporativo
    doc.setFillColor(31, 78, 120);
    doc.rect(0, 0, 297, 18, 'F');
    doc.setFontSize(14);
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.text('FERRETERÍA S.A. - SISTEMA DE REPORTES', 14, 12);

    // Título y fecha del reporte
    doc.setFontSize(14);
    doc.setTextColor(40, 40, 40);
    doc.text(tituloReporte, 14, 28);

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 100, 100);
    doc.text(`Fecha de generación: ${fechaActual}`, 14, 34);

    return doc;
  };

  // Pie de página dinámico "Página X de Y"
  const agregarPiePaginaPDF = (doc: jsPDF) => {
    const totalPages = doc.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFontSize(9);
      doc.setTextColor(120, 120, 120);
      doc.text(
        `Página ${i} de ${totalPages}`,
        doc.internal.pageSize.width - 30,
        doc.internal.pageSize.height - 10
      );
    }
  };

  // ----------------------------------------------------
  // REPORTE 1: Artículos por Nombre (Dialog HTML5)
  // ----------------------------------------------------
  const handleBuscarArticulos = async () => {
    try {
      const data = await getArticulosFiltroNombre(filtroNombre);
      setArticulosFiltrados(data);
    } catch (error) {
      console.error('Error al buscar artículos:', error);
    }
  };

  const exportarReporte1PDF = () => {
    const doc = crearPDFBase('Reporte de Artículos Filtrados por Nombre');

    autoTable(doc, {
      startY: 40,
      head: [['ID', 'Nombre', 'Marca', 'Precio (Q)', 'Stock']],
      body: articulosFiltrados.map((a: Articulo) => [
        a.idArticulo,
        a.nombre,
        a.marca,
        `Q ${Number(a.precio).toFixed(2)}`,
        a.stock
      ]),
      headStyles: { fillColor: [31, 78, 120], textColor: [255, 255, 255], fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [245, 247, 250] },
      styles: { fontSize: 9, cellPadding: 3 }
    });

    agregarPiePaginaPDF(doc);
    doc.save('Reporte_Articulos_Filtrados.pdf');
  };

  const exportarReporte1Excel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Artículos Filtrados');

    worksheet.mergeCells('A1:E1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = 'REPORTE DE ARTÍCULOS FILTRADOS POR NOMBRE';
    titleCell.font = { bold: true, size: 14, color: { argb: '1F4E78' } };

    worksheet.addRow([]);
    worksheet.addRow(['ID', 'Nombre', 'Marca', 'Precio', 'Stock']);

    const headerRow = worksheet.getRow(3);
    headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
    });

    articulosFiltrados.forEach((a: Articulo) => {
      worksheet.addRow([a.idArticulo, a.nombre, a.marca, Number(a.precio), a.stock]);
    });

    worksheet.getColumn(4).numFmt = '"Q"#,##0.00';
    worksheet.columns.forEach((col) => (col.width = 20));

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Reporte_Articulos_Filtrados.xlsx');
  };

  // ----------------------------------------------------
  // REPORTE 2: Artículos por Categoría
  // ----------------------------------------------------
  const exportarReporte2PDF = async () => {
    try {
      const data: ArticuloCategoriaDTO[] = await getReporteArticuloCategoria();
      const doc = crearPDFBase('Reporte de Artículos con su Categoría');

      autoTable(doc, {
        startY: 40,
        head: [['Categoría', 'Artículo', 'Marca', 'Precio (Q)', 'Stock']],
        body: data.map((item: ArticuloCategoriaDTO) => [
          item.nombreCategoria,
          item.nombreArticulo,
          item.marca,
          `Q ${Number(item.precio).toFixed(2)}`,
          item.stock
        ]),
        headStyles: { fillColor: [40, 116, 166], textColor: [255, 255, 255], fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [240, 248, 255] },
        styles: { fontSize: 9, cellPadding: 3 }
      });

      agregarPiePaginaPDF(doc);
      doc.save('Reporte_Articulos_Por_Categoria.pdf');
    } catch (error) {
      console.error('Error al exportar PDF:', error);
    }
  };

  const exportarReporte2Excel = async () => {
    try {
      const data: ArticuloCategoriaDTO[] = await getReporteArticuloCategoria();
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Artículos por Categoría');

      worksheet.mergeCells('A1:E1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = 'REPORTE DE ARTÍCULOS CON SU CATEGORÍA';
      titleCell.font = { bold: true, size: 14, color: { argb: '2874A6' } };

      worksheet.addRow([]);
      worksheet.addRow(['Categoría', 'Artículo', 'Marca', 'Precio', 'Stock']);

      const headerRow = worksheet.getRow(3);
      headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
      headerRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '2874A6' } };
      });

      data.forEach((item: ArticuloCategoriaDTO) => {
        worksheet.addRow([
          item.nombreCategoria,
          item.nombreArticulo,
          item.marca,
          Number(item.precio),
          item.stock
        ]);
      });

      worksheet.getColumn(4).numFmt = '"Q"#,##0.00';
      worksheet.columns.forEach((col) => (col.width = 22));

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), 'Reporte_Articulos_Por_Categoria.xlsx');
    } catch (error) {
      console.error('Error al exportar Excel:', error);
    }
  };

  // ----------------------------------------------------
  // REPORTE 3: Facturas por Cliente
  // ----------------------------------------------------
  const exportarReporte3PDF = async () => {
    try {
      const data: ClienteFacturaDTO[] = await getReporteClienteFactura();
      const doc = crearPDFBase('Reporte de Facturas por Cliente');

      autoTable(doc, {
        startY: 40,
        head: [['Cliente', 'NIT', 'Fecha Factura', 'Total (Q)']],
        body: data.map((item: ClienteFacturaDTO) => [
          `${item.nombreCliente} ${item.apellidoCliente}`,
          item.nit,
          new Date(item.fechaFactura).toLocaleString('es-GT'),
          `Q ${Number(item.totalFactura).toFixed(2)}`
        ]),
        headStyles: { fillColor: [120, 40, 140], textColor: [255, 255, 255], fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [248, 240, 255] },
        styles: { fontSize: 9, cellPadding: 3 }
      });

      agregarPiePaginaPDF(doc);
      doc.save('Reporte_Facturas_Por_Cliente.pdf');
    } catch (error) {
      console.error('Error al exportar PDF:', error);
    }
  };

  const exportarReporte3Excel = async () => {
    try {
      const data: ClienteFacturaDTO[] = await getReporteClienteFactura();
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Facturas por Cliente');

      worksheet.mergeCells('A1:D1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = 'REPORTE DE FACTURAS POR CLIENTE';
      titleCell.font = { bold: true, size: 14, color: { argb: '78288C' } };

      worksheet.addRow([]);
      worksheet.addRow(['Cliente', 'NIT', 'Fecha Factura', 'Total']);

      const headerRow = worksheet.getRow(3);
      headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
      headerRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '78288C' } };
      });

      data.forEach((item: ClienteFacturaDTO) => {
        worksheet.addRow([
          `${item.nombreCliente} ${item.apellidoCliente}`,
          item.nit,
          new Date(item.fechaFactura).toLocaleDateString('es-GT'),
          Number(item.totalFactura)
        ]);
      });

      worksheet.getColumn(4).numFmt = '"Q"#,##0.00';
      worksheet.columns.forEach((col) => (col.width = 25));

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), 'Reporte_Facturas_Por_Cliente.xlsx');
    } catch (error) {
      console.error('Error al exportar Excel:', error);
    }
  };

  return (
    <div style={{ padding: '2rem', fontFamily: 'Arial, sans-serif', maxWidth: '1000px', margin: '0 auto' }}>
      <h2>Módulo de Reportes de Ferretería</h2>
      <p style={{ color: '#666' }}>Generación e impresión de reportes en PDF y Excel</p>
      <hr style={{ marginBottom: '2rem' }} />

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '1.5rem' }}>
        
        {/* Tarjeta 1 */}
        <div style={{ border: '1px solid #ddd', padding: '1.5rem', borderRadius: '8px', backgroundColor: '#fff', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <h3 style={{ marginTop: 0, color: '#1F4E78' }}>1. Artículos por Nombre</h3>
          <p style={{ fontSize: '0.9rem', color: '#555' }}>Filtrar artículos por nombre y generar exportaciones en ventana emergente.</p>
          <button
            onClick={() => dialogRef.current?.showModal()}
            style={{ padding: '0.6rem 1.2rem', backgroundColor: '#1F4E78', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
          >
            Abrir Filtro y Generar
          </button>
        </div>

        {/* Tarjeta 2 */}
        <div style={{ border: '1px solid #ddd', padding: '1.5rem', borderRadius: '8px', backgroundColor: '#fff', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <h3 style={{ marginTop: 0, color: '#2874A6' }}>2. Artículos por Categoría</h3>
          <p style={{ fontSize: '0.9rem', color: '#555' }}>Listado general de artículos ordenados por su categoría correspondiente.</p>
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <button
              onClick={exportarReporte2PDF}
              style={{ padding: '0.5rem 1rem', backgroundColor: '#e74c3c', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
            >
              Exportar PDF
            </button>
            <button
              onClick={exportarReporte2Excel}
              style={{ padding: '0.5rem 1rem', backgroundColor: '#27ae60', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
            >
              Exportar Excel
            </button>
          </div>
        </div>

        {/* Tarjeta 3 */}
        <div style={{ border: '1px solid #ddd', padding: '1.5rem', borderRadius: '8px', backgroundColor: '#fff', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
          <h3 style={{ marginTop: 0, color: '#78288C' }}>3. Facturas por Cliente</h3>
          <p style={{ fontSize: '0.9rem', color: '#555' }}>Reporte de compras por cliente con fecha y totales formateados.</p>
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <button
              onClick={exportarReporte3PDF}
              style={{ padding: '0.5rem 1rem', backgroundColor: '#e74c3c', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
            >
              Exportar PDF
            </button>
            <button
              onClick={exportarReporte3Excel}
              style={{ padding: '0.5rem 1rem', backgroundColor: '#27ae60', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
            >
              Exportar Excel
            </button>
          </div>
        </div>

      </div>

      {/* Dialog HTML5 (<dialog>) para Reporte 1 */}
      <dialog
        ref={dialogRef}
        style={{
          padding: '1.5rem',
          borderRadius: '8px',
          border: '1px solid #ccc',
          maxWidth: '500px',
          width: '90%',
          boxShadow: '0 4px 20px rgba(0,0,0,0.15)'
        }}
      >
        <h3 style={{ marginTop: 0, color: '#1F4E78' }}>Filtrar Artículos por Nombre</h3>
        <div style={{ marginBottom: '1rem' }}>
          <label style={{ display: 'block', marginBottom: '0.5rem', fontSize: '0.9rem' }}>Nombre del artículo:</label>
          <input
            type="text"
            placeholder="Ej. Martillo, Taladro..."
            value={filtroNombre}
            onChange={(e: React.ChangeEvent<HTMLInputElement>) => setFiltroNombre(e.target.value)}
            style={{ width: '100%', padding: '0.5rem', boxSizing: 'border-box', borderRadius: '4px', border: '1px solid #ccc' }}
          />
          <button
            onClick={handleBuscarArticulos}
            style={{ marginTop: '0.5rem', padding: '0.4rem 1rem', backgroundColor: '#34495e', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
          >
            Buscar
          </button>
        </div>

        {articulosFiltrados.length > 0 ? (
          <p style={{ color: '#27ae60', fontWeight: 'bold' }}>Se encontraron {articulosFiltrados.length} registro(s).</p>
        ) : (
          <p style={{ color: '#7f8c8d', fontSize: '0.9rem' }}>Ingrese un término y haga clic en Buscar.</p>
        )}

        <div style={{ marginTop: '1.5rem', display: 'flex', gap: '0.5rem', justifyContent: 'flex-end' }}>
          <button
            onClick={exportarReporte1PDF}
            disabled={articulosFiltrados.length === 0}
            style={{
              padding: '0.5rem 1rem',
              backgroundColor: articulosFiltrados.length === 0 ? '#bdc3c7' : '#e74c3c',
              color: '#fff',
              border: 'none',
              borderRadius: '4px',
              cursor: articulosFiltrados.length === 0 ? 'not-allowed' : 'pointer'
            }}
          >
            Exportar PDF
          </button>
          <button
            onClick={exportarReporte1Excel}
            disabled={articulosFiltrados.length === 0}
            style={{
              padding: '0.5rem 1rem',
              backgroundColor: articulosFiltrados.length === 0 ? '#bdc3c7' : '#27ae60',
              color: '#fff',
              border: 'none',
              borderRadius: '4px',
              cursor: articulosFiltrados.length === 0 ? 'not-allowed' : 'pointer'
            }}
          >
            Exportar Excel
          </button>
          <button
            onClick={() => dialogRef.current?.close()}
            style={{ padding: '0.5rem 1rem', backgroundColor: '#95a5a6', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
          >
            Cerrar
          </button>
        </div>
      </dialog>
    </div>
  );
};