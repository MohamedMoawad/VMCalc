import { createRoot } from 'react-dom/client';
import { StrictMode } from 'react';
import { CostSheet } from '../src/cost/CostSheet';
import { BrowserStore } from './browserStore';

const root = document.getElementById('root');
if (!root) {
  throw new Error('Missing #root');
}

createRoot(root).render(
  <StrictMode>
    <CostSheet store={new BrowserStore('QAR')} userName="You" theme="light" />
  </StrictMode>
);
