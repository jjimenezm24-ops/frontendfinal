export interface Articulo {
  idArticulo: number;
  nombre: string;
  marca: string;
  precio: number;
  stock: number;
  estado: boolean;
}

export interface ArticuloCategoriaDTO {
  nombreArticulo: string;
  marca: string;
  nombreCategoria: string;
  precio: number;
  stock: number;
}

export interface ClienteFacturaDTO {
  nombreCliente: string;
  apellidoCliente: string;
  nit: string;
  fechaFactura: string;
  totalFactura: number;
}