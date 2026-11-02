import { Link, Route, Routes } from 'react-router-dom';
import { useEffect, useState } from 'react';

function DetailsPage() {
  const [form, setForm] = useState({
    name: '',
    yearDept: '',
    collegeSchool: '',
    email: '',
    address: '',
    mobile1: '',
    mobile2: '',
    pincode: '',
    title: '',
    howKnow: '',
    signature: '',
    feesFixed: '',
    dateJoining: '',
    certNo: '',
    dateIssue: '',
    term1: false,
    term2: false
  });
  const [message, setMessage] = useState('');

  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setForm((prev) => ({
      ...prev,
      [name]: type === 'checkbox' ? checked : value
    }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    const res = await fetch('/api/details', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(form)
    });
    const data = await res.json();
    setMessage(data.success ? 'Details saved successfully!' : 'Failed to save details');
  };

  return (
    <div className="page">
      <h1>Saved User Details</h1>
      <p>React + Node full-stack form for Indian Info Tech.</p>
      {message && <p className="msg">{message}</p>}
      <form onSubmit={handleSubmit}>
        <input name="name" placeholder="Name" value={form.name} onChange={handleChange} />
        <input name="yearDept" placeholder="Year / Dept" value={form.yearDept} onChange={handleChange} />
        <input name="collegeSchool" placeholder="College / School" value={form.collegeSchool} onChange={handleChange} />
        <input name="email" placeholder="Email" value={form.email} onChange={handleChange} />
        <textarea name="address" placeholder="Address" value={form.address} onChange={handleChange} />
        <input name="mobile1" placeholder="Mobile 1" value={form.mobile1} onChange={handleChange} />
        <input name="mobile2" placeholder="Mobile 2" value={form.mobile2} onChange={handleChange} />
        <input name="pincode" placeholder="Pincode" value={form.pincode} onChange={handleChange} />
        <input name="title" placeholder="Title Course / Project" value={form.title} onChange={handleChange} />
        <input name="howKnow" placeholder="How you know about this institute" value={form.howKnow} onChange={handleChange} />
        <input name="signature" placeholder="Signature" value={form.signature} onChange={handleChange} />
        <input name="feesFixed" placeholder="Fees Fixed" value={form.feesFixed} onChange={handleChange} />
        <input name="dateJoining" type="date" value={form.dateJoining} onChange={handleChange} />
        <input name="certNo" placeholder="Certificate No" value={form.certNo} onChange={handleChange} />
        <input name="dateIssue" type="date" value={form.dateIssue} onChange={handleChange} />
        <label><input name="term1" type="checkbox" checked={form.term1} onChange={handleChange} /> Term 1</label>
        <label><input name="term2" type="checkbox" checked={form.term2} onChange={handleChange} /> Term 2</label>
        <button type="submit">Save</button>
      </form>
    </div>
  );
}

function SubmissionsPage() {
  const [items, setItems] = useState([]);

  useEffect(() => {
    fetch('/api/submissions')
      .then((res) => res.json())
      .then((data) => setItems(data));
  }, []);

  return (
    <div className="page">
      <h1>Saved Submissions</h1>
      <Link to="/">Back to Details</Link>
      {items.length === 0 ? <p>No submissions yet.</p> : items.map((item, index) => (
        <div key={index} className="card">
          <h3>{item.name || 'Unnamed'}</h3>
          <p>{item.email}</p>
          <p>{item.submittedAt}</p>
        </div>
      ))}
    </div>
  );
}

export default function App() {
  return (
    <div>
      <nav>
        <Link to="/">Details</Link>
        <Link to="/submissions">Submissions</Link>
      </nav>
      <Routes>
        <Route path="/" element={<DetailsPage />} />
        <Route path="/submissions" element={<SubmissionsPage />} />
      </Routes>
    </div>
  );
}
