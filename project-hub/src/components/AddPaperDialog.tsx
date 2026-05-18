import { useState, useEffect } from 'react';
import { useDashboard } from '../context/DashboardContext';
import type { Paper, PaperStatus, Author, Deliverable, Deadline, Link, CoauthorTask, TimelineEvent } from '../types/paper';

const inputStyle: React.CSSProperties = {
  width: '100%',
  padding: '8px 10px',
  border: '1px solid #d0ccc0',
  borderRadius: 4,
  fontSize: 13,
  background: '#faf8f3',
  fontFamily: 'inherit',
  boxSizing: 'border-box',
};

const smallInputStyle: React.CSSProperties = {
  ...inputStyle,
  padding: '5px 8px',
  fontSize: 11,
};

const labelStyle: React.CSSProperties = {
  display: 'block',
  fontSize: 11,
  fontWeight: 700,
  textTransform: 'uppercase',
  letterSpacing: 0.5,
  color: '#888',
  marginBottom: 4,
};

const sectionHeaderStyle: React.CSSProperties = {
  fontSize: 11,
  fontWeight: 700,
  textTransform: 'uppercase',
  color: '#2a7a8a',
  letterSpacing: 0.5,
  marginBottom: 8,
  marginTop: 16,
  padding: '6px 0',
  borderBottom: '1px solid #e8e4da',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
};

const addBtnStyle: React.CSSProperties = {
  background: 'none',
  border: '1px dashed #ccc',
  borderRadius: 4,
  color: '#888',
  cursor: 'pointer',
  fontSize: 10,
  padding: '3px 8px',
};

const removeBtnStyle: React.CSSProperties = {
  background: 'none',
  border: 'none',
  color: '#ccc',
  cursor: 'pointer',
  fontSize: 14,
  padding: '0 4px',
};

function createEmptyPaper(): Paper {
  return {
    id: `project-${Date.now()}`,
    title: '',
    shortTitle: '',
    authors: [],
    progress: { status: 'planning', journal: '', timeline: [] },
    submissionHistory: [],
    links: [],
    deliverables: [],
    deadlines: [],
    notes: [],
    todos: [],
    statistics: { wordCount: 0, figureCount: 0, tableCount: 0, referenceCount: 0 },
    costs: { apcPaid: 0, fundingSources: [] },
    coauthorTasks: [],
  };
}

export function AddPaperDialog() {
  const { addPaper, updatePaper, setAddPaperDialogOpen, editingPaper, setEditingPaper } = useDashboard();
  const isEditing = editingPaper !== null;

  // Basic fields
  const [title, setTitle] = useState('');
  const [shortTitle, setShortTitle] = useState('');
  const [status, setStatus] = useState<PaperStatus>('planning');
  const [journal, setJournal] = useState('');
  const [impactFactor, setImpactFactor] = useState('');
  const [submissionDate, setSubmissionDate] = useState('');
  const [revisionDue, setRevisionDue] = useState('');
  const [repoUrl, setRepoUrl] = useState('');
  const [portalUrl, setPortalUrl] = useState('');
  const [notes, setNotes] = useState('');
  const [wordLimit, setWordLimit] = useState('');
  const [figureLimit, setFigureLimit] = useState('');
  const [tableLimit, setTableLimit] = useState('');
  const [refLimit, setRefLimit] = useState('');

  // Authors (expanded)
  const [authors, setAuthors] = useState<Author[]>([]);

  // Timeline events
  const [timelineEvents, setTimelineEvents] = useState<TimelineEvent[]>([]);

  // Deliverables
  const [deliverables, setDeliverables] = useState<Deliverable[]>([]);

  // Deadlines
  const [deadlines, setDeadlines] = useState<Deadline[]>([]);

  // Extra links
  const [extraLinks, setExtraLinks] = useState<Link[]>([]);

  // Costs
  const [apcEstimate, setApcEstimate] = useState('');
  const [apcCurrency, setApcCurrency] = useState('USD');

  // Coauthor tasks
  const [coauthorTasks, setCoauthorTasks] = useState<CoauthorTask[]>([]);

  // Section toggles
  const [expandedSections, setExpandedSections] = useState<Record<string, boolean>>({});

  const toggleSection = (key: string) => setExpandedSections((prev) => ({ ...prev, [key]: !prev[key] }));

  useEffect(() => {
    if (editingPaper) {
      /* eslint-disable react-hooks/set-state-in-effect */
      setTitle(editingPaper.title);
      setShortTitle(editingPaper.shortTitle);
      setStatus(editingPaper.progress.status);
      setJournal(editingPaper.progress.journal);
      setImpactFactor(editingPaper.progress.impactFactor?.toString() ?? '');
      setSubmissionDate(editingPaper.progress.submissionDate ?? '');
      setRevisionDue(editingPaper.progress.revisionDueDate ?? '');
      setRepoUrl(editingPaper.links.find((l) => l.category === 'repository')?.url ?? '');
      setPortalUrl(editingPaper.links.find((l) => l.category === 'board')?.url ?? '');
      setNotes(editingPaper.notes.map((n) => n.content).join('\n'));
      setWordLimit(editingPaper.statistics.wordLimit?.toString() ?? '');
      setFigureLimit(editingPaper.statistics.figureLimit?.toString() ?? '');
      setTableLimit(editingPaper.statistics.tableLimit?.toString() ?? '');
      setRefLimit(editingPaper.statistics.referenceLimit?.toString() ?? '');
      setAuthors([...editingPaper.authors]);
      setTimelineEvents([...editingPaper.progress.timeline]);
      setDeliverables([...editingPaper.deliverables]);
      setDeadlines([...editingPaper.deadlines]);
      setExtraLinks(editingPaper.links.filter((l) => l.category !== 'repository' && l.category !== 'board'));
      setApcEstimate(editingPaper.costs.apcEstimate?.toString() ?? '');
      setApcCurrency(editingPaper.costs.apcCurrency ?? 'USD');
      setCoauthorTasks([...editingPaper.coauthorTasks]);
      /* eslint-enable react-hooks/set-state-in-effect */
    }
  }, [editingPaper]);

  const handleClose = () => {
    setAddPaperDialogOpen(false);
    setEditingPaper(null);
  };

  const handleSave = () => {
    const links: Link[] = [];
    if (portalUrl) links.push({ label: `${journal || 'Project'} Board`, url: portalUrl, category: 'board' });
    if (repoUrl) links.push({ label: repoUrl.split('/').slice(-1)[0] || 'Repository', url: repoUrl, category: 'repository' });
    links.push(...extraLinks);

    const timeline = [...timelineEvents];
    if (submissionDate && !timeline.some((t) => t.type === 'planning')) {
      timeline.push({ label: 'Kickoff', startDate: submissionDate, endDate: submissionDate, type: 'planning' });
    }

    const paper: Paper = isEditing
      ? {
          ...editingPaper,
          title,
          shortTitle: shortTitle || title.slice(0, 30),
          authors: authors.length > 0 ? authors : editingPaper.authors,
          progress: {
            ...editingPaper.progress,
            status,
            journal,
            impactFactor: impactFactor ? parseFloat(impactFactor) : undefined,
            submissionDate: submissionDate || editingPaper.progress.submissionDate,
            revisionDueDate: revisionDue || editingPaper.progress.revisionDueDate,
            timeline: timeline.length > 0 ? timeline : editingPaper.progress.timeline,
          },
          submissionHistory: editingPaper.submissionHistory ?? [],
          links: links.length > 0 ? links : editingPaper.links,
          deliverables: deliverables.length > 0 ? deliverables : editingPaper.deliverables,
          deadlines: deadlines.length > 0 ? deadlines : editingPaper.deadlines,
          statistics: {
            ...editingPaper.statistics,
            wordLimit: wordLimit ? parseInt(wordLimit) : editingPaper.statistics.wordLimit,
            figureLimit: figureLimit ? parseInt(figureLimit) : editingPaper.statistics.figureLimit,
            tableLimit: tableLimit ? parseInt(tableLimit) : editingPaper.statistics.tableLimit,
            referenceLimit: refLimit ? parseInt(refLimit) : editingPaper.statistics.referenceLimit,
          },
          costs: {
            ...editingPaper.costs,
            apcEstimate: apcEstimate ? parseFloat(apcEstimate) : editingPaper.costs.apcEstimate,
            apcCurrency: apcCurrency,
          },
          coauthorTasks: coauthorTasks.length > 0 ? coauthorTasks : editingPaper.coauthorTasks,
        }
      : {
          ...createEmptyPaper(),
          title,
          shortTitle: shortTitle || title.slice(0, 30),
          authors,
          progress: {
            status,
            journal,
            impactFactor: impactFactor ? parseFloat(impactFactor) : undefined,
            submissionDate: submissionDate || undefined,
            revisionDueDate: revisionDue || undefined,
            timeline,
          },
          submissionHistory: [],
          links,
          deliverables,
          deadlines,
          notes: notes ? [{ id: `n-${Date.now()}`, content: notes, createdAt: new Date().toISOString().slice(0, 10) }] : [],
          statistics: {
            wordCount: 0,
            wordLimit: wordLimit ? parseInt(wordLimit) : undefined,
            figureCount: 0,
            figureLimit: figureLimit ? parseInt(figureLimit) : undefined,
            tableCount: 0,
            tableLimit: tableLimit ? parseInt(tableLimit) : undefined,
            referenceCount: 0,
            referenceLimit: refLimit ? parseInt(refLimit) : undefined,
          },
          costs: {
            apcEstimate: apcEstimate ? parseFloat(apcEstimate) : undefined,
            apcCurrency,
            apcPaid: 0,
            fundingSources: [],
          },
          coauthorTasks,
        };

    if (isEditing) {
      updatePaper(paper);
    } else {
      addPaper(paper);
    }
    handleClose();
  };

  // Author helpers
  const addAuthor = () => setAuthors([...authors, { name: '', role: 'member' }]);
  const updateAuthorField = (i: number, field: keyof Author, val: string) =>
    setAuthors(authors.map((a, j) => (j === i ? { ...a, [field]: val || undefined } : a)));
  const removeAuthor = (i: number) => setAuthors(authors.filter((_, j) => j !== i));

  // Timeline event helpers
  const addTimelineEvent = () => setTimelineEvents([...timelineEvents, { label: '', startDate: '', type: 'milestone' }]);
  const updateEventField = (i: number, field: string, val: string) =>
    setTimelineEvents(timelineEvents.map((e, j) => (j === i ? { ...e, [field]: val || undefined } : e)));
  const removeEvent = (i: number) => setTimelineEvents(timelineEvents.filter((_, j) => j !== i));

  // Deliverable helpers
  const addDeliverable = () => setDeliverables([...deliverables, { name: '', type: 'feature', status: 'pending' }]);
  const updateDeliverableField = (i: number, field: string, val: string) =>
    setDeliverables(deliverables.map((d, j) => (j === i ? { ...d, [field]: val || undefined } : d)));
  const removeDeliverable = (i: number) => setDeliverables(deliverables.filter((_, j) => j !== i));

  // Deadline helpers
  const addDeadline = () => setDeadlines([...deadlines, { label: '', date: '', type: 'other' }]);
  const updateDeadlineField = (i: number, field: string, val: string) =>
    setDeadlines(deadlines.map((d, j) => (j === i ? { ...d, [field]: val } : d)));
  const removeDeadline = (i: number) => setDeadlines(deadlines.filter((_, j) => j !== i));

  // Extra link helpers
  const addExtraLink = () => setExtraLinks([...extraLinks, { label: '', url: '', category: 'other' }]);
  const updateLinkField = (i: number, field: string, val: string) =>
    setExtraLinks(extraLinks.map((l, j) => (j === i ? { ...l, [field]: val } : l)));
  const removeExtraLink = (i: number) => setExtraLinks(extraLinks.filter((_, j) => j !== i));

  // Coauthor task helpers
  const addCoauthorTask = () => setCoauthorTasks([...coauthorTasks, { authorName: '', initials: '', task: '', status: 'in-progress' }]);
  const updateTaskField = (i: number, field: string, val: string) => {
    setCoauthorTasks(coauthorTasks.map((t, j) => {
      if (j !== i) return t;
      const updated = { ...t, [field]: val };
      if (field === 'authorName') {
        updated.initials = val.split(/\s+/).map((w) => w[0] || '').join('').toUpperCase().slice(0, 2);
      }
      return updated;
    }));
  };
  const removeCoauthorTask = (i: number) => setCoauthorTasks(coauthorTasks.filter((_, j) => j !== i));

  const isExpanded = (key: string) => expandedSections[key] ?? false;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        background: 'rgba(0,0,0,0.4)',
        zIndex: 500,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
      onClick={(e) => e.target === e.currentTarget && handleClose()}
    >
      <div
        style={{
          background: '#fff',
          borderRadius: 10,
          width: 680,
          maxHeight: '90vh',
          overflowY: 'auto',
          boxShadow: '0 16px 48px rgba(0,0,0,0.2)',
        }}
      >
        <div
          style={{
            padding: '16px 20px',
            borderBottom: '1px solid #e8e4da',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
          }}
        >
          <h2 style={{ fontSize: 16, fontWeight: 700, color: '#1a3a4a', margin: 0 }}>
            {isEditing ? 'Edit Project' : 'Add New Project'}
          </h2>
          <button
            onClick={handleClose}
            style={{ background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#999' }}
          >
            &times;
          </button>
        </div>

        <div style={{ padding: 20 }}>
          {/* Basic Info */}
          <div style={{ marginBottom: 14 }}>
            <label style={labelStyle}>Project Title</label>
            <input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g., Website Redesign, Mobile App v2..." />
          </div>

          <div style={{ display: 'flex', gap: 12, marginBottom: 14 }}>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Short Title (optional)</label>
              <input style={inputStyle} value={shortTitle} onChange={(e) => setShortTitle(e.target.value)} placeholder="Auto-generated if empty" />
            </div>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Status</label>
              <select style={inputStyle} value={status} onChange={(e) => setStatus(e.target.value as PaperStatus)}>
                <option value="planning">Planning</option>
                <option value="in-progress">In Progress</option>
                <option value="in-review">In Review</option>
                <option value="on-hold">On Hold</option>
                <option value="testing">Testing</option>
                <option value="approved">Approved</option>
                <option value="completed">Completed</option>
                <option value="cancelled">Cancelled</option>
              </select>
            </div>
          </div>

          <div style={{ display: 'flex', gap: 12, marginBottom: 14 }}>
            <div style={{ flex: 2 }}>
              <label style={labelStyle}>Client / Category</label>
              <input style={inputStyle} value={journal} onChange={(e) => setJournal(e.target.value)} placeholder="e.g., Acme Corp, Internal, R&D" />
            </div>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Priority (1-10)</label>
              <input style={inputStyle} type="number" step="1" min="1" max="10" value={impactFactor} onChange={(e) => setImpactFactor(e.target.value)} placeholder="e.g., 8" />
            </div>
          </div>

          <div style={{ display: 'flex', gap: 12, marginBottom: 14 }}>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Start Date</label>
              <input style={inputStyle} type="date" value={submissionDate} onChange={(e) => setSubmissionDate(e.target.value)} />
            </div>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Deadline</label>
              <input style={inputStyle} type="date" value={revisionDue} onChange={(e) => setRevisionDue(e.target.value)} />
            </div>
          </div>

          <div style={{ display: 'flex', gap: 12, marginBottom: 14 }}>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Project Board URL</label>
              <input style={inputStyle} value={portalUrl} onChange={(e) => setPortalUrl(e.target.value)} placeholder="https://..." />
            </div>
            <div style={{ flex: 1 }}>
              <label style={labelStyle}>Git Repository URL</label>
              <input style={inputStyle} value={repoUrl} onChange={(e) => setRepoUrl(e.target.value)} placeholder="https://github.com/..." />
            </div>
          </div>

          {/* Authors Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('authors')}>
            <span>Team Members ({authors.length})</span>
            <span style={{ fontSize: 14 }}>{isExpanded('authors') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('authors') && (
            <div style={{ marginBottom: 8 }}>
              {authors.map((a, i) => (
                <div key={i} style={{ display: 'flex', gap: 4, marginBottom: 4, alignItems: 'center' }}>
                  <input style={{ ...smallInputStyle, flex: 2 }} value={a.name} onChange={(e) => updateAuthorField(i, 'name', e.target.value)} placeholder="Name" />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={a.role} onChange={(e) => updateAuthorField(i, 'role', e.target.value)}>
                    <option value="lead">Lead</option>
                    <option value="manager">Manager</option>
                    <option value="member">Member</option>
                  </select>
                  <input style={{ ...smallInputStyle, flex: 2 }} value={a.affiliation ?? ''} onChange={(e) => updateAuthorField(i, 'affiliation', e.target.value)} placeholder="Team / Department" />
                  <input style={{ ...smallInputStyle, flex: 2 }} value={a.email ?? ''} onChange={(e) => updateAuthorField(i, 'email', e.target.value)} placeholder="Email" />
                  <button onClick={() => removeAuthor(i)} style={removeBtnStyle}>&times;</button>
                </div>
              ))}
              <button onClick={addAuthor} style={addBtnStyle}>+ Add Member</button>
            </div>
          )}

          {/* Timeline Events Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('timeline')}>
            <span>Timeline Events ({timelineEvents.length})</span>
            <span style={{ fontSize: 14 }}>{isExpanded('timeline') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('timeline') && (
            <div style={{ marginBottom: 8 }}>
              {timelineEvents.map((ev, i) => (
                <div key={i} style={{ display: 'flex', gap: 4, marginBottom: 4, alignItems: 'center' }}>
                  <input style={{ ...smallInputStyle, flex: 2 }} value={ev.label} onChange={(e) => updateEventField(i, 'label', e.target.value)} placeholder="Label" />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={ev.type} onChange={(e) => updateEventField(i, 'type', e.target.value)}>
                    <option value="planning">Planning</option>
                    <option value="development">Development</option>
                    <option value="review">Review</option>
                    <option value="testing">Testing</option>
                    <option value="deployment">Deployment</option>
                    <option value="milestone">Milestone</option>
                  </select>
                  <input type="date" style={{ ...smallInputStyle, flex: 1 }} value={ev.startDate} onChange={(e) => updateEventField(i, 'startDate', e.target.value)} />
                  <input type="date" style={{ ...smallInputStyle, flex: 1 }} value={ev.endDate ?? ''} onChange={(e) => updateEventField(i, 'endDate', e.target.value)} placeholder="End" />
                  <button onClick={() => removeEvent(i)} style={removeBtnStyle}>&times;</button>
                </div>
              ))}
              <button onClick={addTimelineEvent} style={addBtnStyle}>+ Add Event</button>
            </div>
          )}

          {/* Deliverables Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('deliverables')}>
            <span>Deliverables ({deliverables.length})</span>
            <span style={{ fontSize: 14 }}>{isExpanded('deliverables') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('deliverables') && (
            <div style={{ marginBottom: 8 }}>
              {deliverables.map((d, i) => (
                <div key={i} style={{ display: 'flex', gap: 4, marginBottom: 4, alignItems: 'center' }}>
                  <input style={{ ...smallInputStyle, flex: 2 }} value={d.name} onChange={(e) => updateDeliverableField(i, 'name', e.target.value)} placeholder="Name" />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={d.type} onChange={(e) => updateDeliverableField(i, 'type', e.target.value)}>
                    <option value="feature">Feature</option>
                    <option value="bugfix">Bugfix</option>
                    <option value="docs">Documentation</option>
                    <option value="design">Design</option>
                    <option value="api">API</option>
                    <option value="infrastructure">Infrastructure</option>
                    <option value="other">Other</option>
                  </select>
                  <input style={{ ...smallInputStyle, flex: 1 }} value={d.version ?? ''} onChange={(e) => updateDeliverableField(i, 'version', e.target.value)} placeholder="Version" />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={d.status} onChange={(e) => updateDeliverableField(i, 'status', e.target.value)}>
                    <option value="pending">Pending</option>
                    <option value="in-progress">In Progress</option>
                    <option value="complete">Complete</option>
                  </select>
                  <button onClick={() => removeDeliverable(i)} style={removeBtnStyle}>&times;</button>
                </div>
              ))}
              <button onClick={addDeliverable} style={addBtnStyle}>+ Add Deliverable</button>
            </div>
          )}

          {/* Deadlines Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('deadlines')}>
            <span>Deadlines ({deadlines.length})</span>
            <span style={{ fontSize: 14 }}>{isExpanded('deadlines') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('deadlines') && (
            <div style={{ marginBottom: 8 }}>
              {deadlines.map((d, i) => (
                <div key={i} style={{ display: 'flex', gap: 4, marginBottom: 4, alignItems: 'center' }}>
                  <input style={{ ...smallInputStyle, flex: 2 }} value={d.label} onChange={(e) => updateDeadlineField(i, 'label', e.target.value)} placeholder="Label" />
                  <input type="date" style={{ ...smallInputStyle, flex: 1 }} value={d.date} onChange={(e) => updateDeadlineField(i, 'date', e.target.value)} />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={d.type} onChange={(e) => updateDeadlineField(i, 'type', e.target.value)}>
                    <option value="sprint">Sprint</option>
                    <option value="milestone">Milestone</option>
                    <option value="release">Release</option>
                    <option value="demo">Demo</option>
                    <option value="other">Other</option>
                  </select>
                  <button onClick={() => removeDeadline(i)} style={removeBtnStyle}>&times;</button>
                </div>
              ))}
              <button onClick={addDeadline} style={addBtnStyle}>+ Add Deadline</button>
            </div>
          )}

          {/* Extra Links Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('links')}>
            <span>Additional Links ({extraLinks.length})</span>
            <span style={{ fontSize: 14 }}>{isExpanded('links') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('links') && (
            <div style={{ marginBottom: 8 }}>
              {extraLinks.map((l, i) => (
                <div key={i} style={{ display: 'flex', gap: 4, marginBottom: 4, alignItems: 'center' }}>
                  <input style={{ ...smallInputStyle, flex: 1 }} value={l.label} onChange={(e) => updateLinkField(i, 'label', e.target.value)} placeholder="Label" />
                  <input style={{ ...smallInputStyle, flex: 2 }} value={l.url} onChange={(e) => updateLinkField(i, 'url', e.target.value)} placeholder="URL" />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={l.category} onChange={(e) => updateLinkField(i, 'category', e.target.value)}>
                    <option value="board">Project Board</option>
                    <option value="devin">Devin</option>
                    <option value="repository">Repository</option>
                    <option value="other">Other</option>
                  </select>
                  <button onClick={() => removeExtraLink(i)} style={removeBtnStyle}>&times;</button>
                </div>
              ))}
              <button onClick={addExtraLink} style={addBtnStyle}>+ Add Link</button>
            </div>
          )}

          {/* Costs Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('costs')}>
            <span>Budget & Funding</span>
            <span style={{ fontSize: 14 }}>{isExpanded('costs') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('costs') && (
            <div style={{ display: 'flex', gap: 12, marginBottom: 8 }}>
              <div style={{ flex: 1 }}>
                <label style={{ ...labelStyle, fontSize: 10 }}>Budget Estimate</label>
                <input style={smallInputStyle} type="number" value={apcEstimate} onChange={(e) => setApcEstimate(e.target.value)} placeholder="0" />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ ...labelStyle, fontSize: 10 }}>Currency</label>
                <select style={smallInputStyle} value={apcCurrency} onChange={(e) => setApcCurrency(e.target.value)}>
                  <option value="USD">USD</option>
                  <option value="GBP">GBP</option>
                  <option value="EUR">EUR</option>
                  <option value="JPY">JPY</option>
                </select>
              </div>
            </div>
          )}

          {/* Co-author Tasks Section */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('tasks')}>
            <span>Team Tasks ({coauthorTasks.length})</span>
            <span style={{ fontSize: 14 }}>{isExpanded('tasks') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('tasks') && (
            <div style={{ marginBottom: 8 }}>
              {coauthorTasks.map((t, i) => (
                <div key={i} style={{ display: 'flex', gap: 4, marginBottom: 4, alignItems: 'center' }}>
                  <input style={{ ...smallInputStyle, flex: 1 }} value={t.authorName} onChange={(e) => updateTaskField(i, 'authorName', e.target.value)} placeholder="Member" />
                  <input style={{ ...smallInputStyle, flex: 2 }} value={t.task} onChange={(e) => updateTaskField(i, 'task', e.target.value)} placeholder="Task description" />
                  <select style={{ ...smallInputStyle, flex: 1 }} value={t.status} onChange={(e) => updateTaskField(i, 'status', e.target.value)}>
                    <option value="in-progress">In Progress</option>
                    <option value="waiting">Waiting</option>
                    <option value="done">Done</option>
                  </select>
                  <button onClick={() => removeCoauthorTask(i)} style={removeBtnStyle}>&times;</button>
                </div>
              ))}
              <button onClick={addCoauthorTask} style={addBtnStyle}>+ Add Task</button>
            </div>
          )}

          {/* Journal Limits */}
          <div style={sectionHeaderStyle} onClick={() => toggleSection('limits')}>
            <span>Capacity Limits</span>
            <span style={{ fontSize: 14 }}>{isExpanded('limits') ? '\u25B2' : '\u25BC'}</span>
          </div>
          {isExpanded('limits') && (
            <div style={{ display: 'flex', gap: 12, marginBottom: 8 }}>
              <div style={{ flex: 1 }}>
                <label style={{ ...labelStyle, fontSize: 10 }}>Task Limit</label>
                <input style={smallInputStyle} type="number" value={wordLimit} onChange={(e) => setWordLimit(e.target.value)} placeholder="60" />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ ...labelStyle, fontSize: 10 }}>Issue Limit</label>
                <input style={smallInputStyle} type="number" value={figureLimit} onChange={(e) => setFigureLimit(e.target.value)} placeholder="15" />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ ...labelStyle, fontSize: 10 }}>Milestone Limit</label>
                <input style={smallInputStyle} type="number" value={tableLimit} onChange={(e) => setTableLimit(e.target.value)} placeholder="5" />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ ...labelStyle, fontSize: 10 }}>Sprint Limit</label>
                <input style={smallInputStyle} type="number" value={refLimit} onChange={(e) => setRefLimit(e.target.value)} placeholder="10" />
              </div>
            </div>
          )}

          {/* Notes */}
          {!isEditing && (
            <div style={{ marginTop: 14, marginBottom: 14 }}>
              <label style={labelStyle}>Notes</label>
              <textarea
                style={{ ...inputStyle, resize: 'vertical', minHeight: 60 }}
                value={notes}
                onChange={(e) => setNotes(e.target.value)}
                placeholder="Initial notes about the project..."
              />
            </div>
          )}
        </div>

        <div
          style={{
            padding: '14px 20px',
            borderTop: '1px solid #e8e4da',
            display: 'flex',
            justifyContent: 'flex-end',
            gap: 8,
          }}
        >
          <button
            onClick={handleClose}
            style={{
              padding: '8px 20px',
              background: 'none',
              border: '1px solid #d0ccc0',
              borderRadius: 4,
              fontSize: 13,
              cursor: 'pointer',
              color: '#777',
            }}
          >
            Cancel
          </button>
          <button
            onClick={handleSave}
            disabled={!title.trim()}
            style={{
              padding: '8px 20px',
              background: title.trim() ? '#2a7a8a' : '#ccc',
              border: 'none',
              borderRadius: 4,
              fontSize: 13,
              color: 'white',
              fontWeight: 600,
              cursor: title.trim() ? 'pointer' : 'not-allowed',
            }}
          >
            {isEditing ? 'Save Changes' : 'Add Project'}
          </button>
        </div>
      </div>
    </div>
  );
}
