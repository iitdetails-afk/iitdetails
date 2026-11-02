import React from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter, Link, Route, Routes } from 'react-router-dom';

function App() {
  return (
    <BrowserRouter>
      <nav style={{ display: 'flex', gap: '12px', padding: '16px', background: '#13263d', color: 'white' }}>
        <Link to="/details.html" style={{ color: 'white', textDecoration: 'none' }}>Details</Link>
        <Link to="/form.html" style={{ color: 'white', textDecoration: 'none' }}>Form</Link>
        <Link to="/ss.html" style={{ color: 'white', textDecoration: 'none' }}>Submissions</Link>
      </nav>
      <Routes>
        <Route path="/details.html" element={<div style={{ padding: '20px' }}>Open the details page via the backend at /details.html</div>} />
        <Route path="/form.html" element={<div style={{ padding: '20px' }}>Open the form page via the backend at /form.html</div>} />
        <Route path="/ss.html" element={<div style={{ padding: '20px' }}>Open the submissions page via the backend at /ss.html</div>} />
      </Routes>
    </BrowserRouter>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><App /></React.StrictMode>
);
