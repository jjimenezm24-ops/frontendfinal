import axios from 'axios';
import { Articulo, ArticuloCategoriaDTO, ClienteFacturaDTO } from '../types/reportes';

const API_BASE = 'http://localhost:8080';

export const getArticulosFiltroNombre = async (nombre: string): Promise<Articulo[]> => {
  const response = await axios.get<Articulo[]>(`${API_BASE}/articulos/mostrarActivosFiltroNombre`, {
    params: { nombre }
  });
  return response.data;
};

export const getReporteArticuloCategoria = async (): Promise<ArticuloCategoriaDTO[]> => {
  const response = await axios.get<ArticuloCategoriaDTO[]>(`${API_BASE}/articulos/reporteArticuloCategoria`);
  return response.data;
};

export const getReporteClienteFactura = async (): Promise<ClienteFacturaDTO[]> => {
  const response = await axios.get<ClienteFacturaDTO[]>(`${API_BASE}/clientes/reporteClienteFactura`);
  return response.data;
};