import React from 'react';
import ReactDOM from 'react-dom/client';
import { Reportes } from './pages/Reportes';

const rootElement = document.getElementById('root');

if (rootElement) {
  const root = ReactDOM.createRoot(rootElement as HTMLElement);
  root.render(
    <React.StrictMode>
      <Reportes />
    </React.StrictMode>
  );
}