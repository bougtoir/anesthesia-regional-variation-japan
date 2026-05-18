import { useState } from 'react';
import type { Paper, FundingSource } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const inputStyle: React.CSSProperties = {
  padding: '4px 8px',
  border: '1px solid #d0ccc0',
  borderRadius: 4,
  fontSize: 11,
  background: '#faf8f3',
  fontFamily: 'inherit',
};

interface Props {
  paper: Paper;
}

function formatCurrency(amount: number, currency: string): string {
  const symbols: Record<string, string> = { GBP: '\u00A3', USD: '$', JPY: '\u00A5', EUR: '\u20AC' };
  const sym = symbols[currency] ?? currency + ' ';
  return `${sym}${amount.toLocaleString()}`;
}

export function CostsFundingPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const c = paper.costs;
  const apcCurrency = c.apcCurrency ?? 'USD';
  const outstanding = (c.apcEstimate ?? 0) - c.apcPaid;

  const [editingApc, setEditingApc] = useState(false);
  const [apcEst, setApcEst] = useState(String(c.apcEstimate ?? 0));
  const [apcPaid, setApcPaid] = useState(String(c.apcPaid));
  const [apcCur, setApcCur] = useState(apcCurrency);

  const [showAddFunding, setShowAddFunding] = useState(false);
  const [fsName, setFsName] = useState('');
  const [fsBudget, setFsBudget] = useState('');
  const [fsAllocated, setFsAllocated] = useState('');
  const [fsCurrency, setFsCurrency] = useState('USD');
  const [fsStatus, setFsStatus] = useState<FundingSource['status']>('pending');

  const saveApc = () => {
    updatePaper({
      ...paper,
      costs: {
        ...paper.costs,
        apcEstimate: parseFloat(apcEst) || 0,
        apcPaid: parseFloat(apcPaid) || 0,
        apcCurrency: apcCur,
      },
    });
    setEditingApc(false);
  };

  const addFunding = () => {
    if (!fsName.trim()) return;
    updatePaper({
      ...paper,
      costs: {
        ...paper.costs,
        fundingSources: [
          ...paper.costs.fundingSources,
          {
            name: fsName.trim(),
            budget: parseFloat(fsBudget) || 0,
            allocated: parseFloat(fsAllocated) || 0,
            currency: fsCurrency,
            status: fsStatus,
          },
        ],
      },
    });
    setFsName('');
    setFsBudget('');
    setFsAllocated('');
    setFsCurrency('USD');
    setFsStatus('pending');
    setShowAddFunding(false);
  };

  const removeFunding = (index: number) => {
    updatePaper({
      ...paper,
      costs: {
        ...paper.costs,
        fundingSources: paper.costs.fundingSources.filter((_, i) => i !== index),
      },
    });
  };

  return (
    <div>
      {/* Budget Section */}
      {editingApc ? (
        <div style={{ marginBottom: 12, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: '#888', marginBottom: 4 }}>Project Budget</div>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={apcEst} onChange={(e) => setApcEst(e.target.value)} placeholder="Estimate" type="number" style={{ ...inputStyle, flex: 1 }} />
            <input value={apcPaid} onChange={(e) => setApcPaid(e.target.value)} placeholder="Paid" type="number" style={{ ...inputStyle, flex: 1 }} />
            <select value={apcCur} onChange={(e) => setApcCur(e.target.value)} style={inputStyle}>
              <option value="USD">USD</option>
              <option value="GBP">GBP</option>
              <option value="EUR">EUR</option>
              <option value="JPY">JPY</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4, justifyContent: 'flex-end' }}>
            <button onClick={saveApc} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Save</button>
            <button onClick={() => setEditingApc(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <>
          {(c.apcEstimate ?? 0) > 0 ? (
            <div style={{ display: 'flex', gap: 12, marginBottom: 12, cursor: 'pointer' }} onClick={() => { setApcEst(String(c.apcEstimate ?? 0)); setApcPaid(String(c.apcPaid)); setApcCur(apcCurrency); setEditingApc(true); }} title="Click to edit Budget">
              <div style={{ flex: 1, padding: 10, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da', textAlign: 'center' }}>
                <div style={{ fontSize: 18, fontWeight: 700, color: '#1a3a4a' }}>{formatCurrency(c.apcEstimate ?? 0, apcCurrency)}</div>
                <div style={{ fontSize: 10, color: '#888' }}>Budget</div>
              </div>
              <div style={{ flex: 1, padding: 10, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da', textAlign: 'center' }}>
                <div style={{ fontSize: 18, fontWeight: 700, color: '#40a060' }}>{formatCurrency(c.apcPaid, apcCurrency)}</div>
                <div style={{ fontSize: 10, color: '#888' }}>Paid</div>
              </div>
              <div style={{ flex: 1, padding: 10, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da', textAlign: 'center' }}>
                <div style={{ fontSize: 18, fontWeight: 700, color: outstanding > 0 ? '#d04040' : '#40a060' }}>{formatCurrency(outstanding, apcCurrency)}</div>
                <div style={{ fontSize: 10, color: '#888' }}>Outstanding</div>
              </div>
            </div>
          ) : (
            <button
              onClick={() => setEditingApc(true)}
              style={{ marginBottom: 8, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
            >
              + Set Budget
            </button>
          )}
        </>
      )}

      {/* Funding Sources Table */}
      {c.fundingSources.length > 0 ? (
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr>
              {['Funding Source', 'Budget', 'Allocated', 'Status', ''].map((h) => (
                <th
                  key={h}
                  style={{
                    textAlign: 'left',
                    fontSize: 10,
                    fontWeight: 700,
                    textTransform: 'uppercase',
                    letterSpacing: 0.5,
                    color: '#888',
                    padding: '4px 8px 6px',
                    borderBottom: '2px solid #d0ccc0',
                  }}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {c.fundingSources.map((fs, i) => (
              <tr key={i}>
                <td style={{ padding: '7px 8px', fontSize: 12, borderBottom: '1px solid #f0ede6', background: i % 2 === 1 ? '#faf8f3' : undefined }}>{fs.name}</td>
                <td style={{ padding: '7px 8px', fontSize: 12, borderBottom: '1px solid #f0ede6', background: i % 2 === 1 ? '#faf8f3' : undefined }}>{formatCurrency(fs.budget, fs.currency)}</td>
                <td style={{ padding: '7px 8px', fontSize: 12, borderBottom: '1px solid #f0ede6', background: i % 2 === 1 ? '#faf8f3' : undefined }}>{formatCurrency(fs.allocated, fs.currency)}</td>
                <td style={{ padding: '7px 8px', fontSize: 12, borderBottom: '1px solid #f0ede6', background: i % 2 === 1 ? '#faf8f3' : undefined }}>
                  <span style={{ color: fs.status === 'approved' ? '#40a060' : fs.status === 'pending' ? '#d09030' : '#c04040', fontWeight: 600 }}>
                    {fs.status.charAt(0).toUpperCase() + fs.status.slice(1)}
                  </span>
                </td>
                <td style={{ padding: '7px 8px', fontSize: 12, borderBottom: '1px solid #f0ede6', background: i % 2 === 1 ? '#faf8f3' : undefined }}>
                  <button onClick={() => removeFunding(i)} style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 12 }} title="Remove">&times;</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      ) : (
        !showAddFunding && <div style={{ color: '#aaa', fontSize: 12 }}>No funding sources configured</div>
      )}

      {/* Add Funding */}
      {showAddFunding ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={fsName} onChange={(e) => setFsName(e.target.value)} placeholder="Source name" style={{ ...inputStyle, flex: 1 }} />
            <select value={fsCurrency} onChange={(e) => setFsCurrency(e.target.value)} style={inputStyle}>
              <option value="JPY">JPY</option>
              <option value="USD">USD</option>
              <option value="GBP">GBP</option>
              <option value="EUR">EUR</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={fsBudget} onChange={(e) => setFsBudget(e.target.value)} placeholder="Budget" type="number" style={{ ...inputStyle, flex: 1 }} />
            <input value={fsAllocated} onChange={(e) => setFsAllocated(e.target.value)} placeholder="Allocated" type="number" style={{ ...inputStyle, flex: 1 }} />
            <select value={fsStatus} onChange={(e) => setFsStatus(e.target.value as FundingSource['status'])} style={inputStyle}>
              <option value="pending">Pending</option>
              <option value="approved">Approved</option>
              <option value="rejected">Rejected</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4, justifyContent: 'flex-end' }}>
            <button onClick={addFunding} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAddFunding(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowAddFunding(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Funding Source
        </button>
      )}
    </div>
  );
}
