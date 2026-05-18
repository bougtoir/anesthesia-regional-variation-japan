import type { TimelineEvent } from '../../types/paper';

const BAR_COLORS: Record<string, string> = {
  drafting: 'linear-gradient(90deg, #e0a040, #e8b860)',
  submission: 'linear-gradient(90deg, #5090d0, #70a8e0)',
  review: 'linear-gradient(90deg, #9070c0, #a888d0)',
  revision: 'linear-gradient(90deg, #d08050, #e09868)',
  decision: 'linear-gradient(90deg, #40a060, #58b878)',
  milestone: 'linear-gradient(90deg, #6090a0, #80b0c0)',
};

interface Props {
  events: TimelineEvent[];
}

export function GanttChart({ events }: Props) {
  if (events.length === 0) return <div style={{ color: '#aaa', fontSize: 12 }}>No timeline events</div>;

  const allDates = events.flatMap((e) => {
    const dates = [new Date(e.startDate)];
    if (e.endDate) dates.push(new Date(e.endDate));
    return dates;
  });
  const now = new Date();
  allDates.push(now);

  const minDate = new Date(Math.min(...allDates.map((d) => d.getTime())));
  const maxDate = new Date(Math.max(...allDates.map((d) => d.getTime())));

  minDate.setDate(1);
  maxDate.setMonth(maxDate.getMonth() + 2);
  maxDate.setDate(1);

  const totalMs = maxDate.getTime() - minDate.getTime();
  const toPercent = (d: Date) => ((d.getTime() - minDate.getTime()) / totalMs) * 100;

  const months: { label: string; left: number }[] = [];
  const cursor = new Date(minDate);
  while (cursor < maxDate) {
    months.push({
      label: cursor.toLocaleString('en', { month: 'short' }),
      left: toPercent(new Date(cursor)),
    });
    cursor.setMonth(cursor.getMonth() + 1);
  }

  const nowPercent = toPercent(now);

  return (
    <div style={{ position: 'relative', marginTop: 8 }}>
      <div style={{ display: 'flex', borderBottom: '2px solid #d0ccc0', marginBottom: 4 }}>
        <div style={{ width: 110, flexShrink: 0 }} />
        <div style={{ flex: 1, display: 'flex', position: 'relative' }}>
          {months.map((m, i) => (
            <div
              key={i}
              style={{
                flex: 1,
                textAlign: 'center',
                fontSize: 10,
                fontWeight: 600,
                color: '#888',
                padding: '4px 0',
              }}
            >
              {m.label}
            </div>
          ))}
        </div>
      </div>

      <div style={{ position: 'relative' }}>
        {events.map((ev, i) => {
          const start = toPercent(new Date(ev.startDate));
          const end = ev.endDate != null ? toPercent(new Date(ev.endDate)) : toPercent(now);
          const width = Math.max(end - start, 2);

          return (
            <div
              key={i}
              style={{
                display: 'flex',
                alignItems: 'center',
                height: 28,
                borderBottom: '1px solid #f0ede6',
                background: i % 2 === 1 ? '#faf8f3' : undefined,
              }}
            >
              <div
                style={{
                  width: 110,
                  flexShrink: 0,
                  fontSize: 11,
                  color: '#555',
                  padding: '0 8px',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                }}
              >
                {ev.label}
              </div>
              <div style={{ flex: 1, position: 'relative', height: '100%', display: 'flex', alignItems: 'center' }}>
                <div
                  style={{
                    position: 'absolute',
                    left: `${start}%`,
                    width: `${width}%`,
                    height: 16,
                    borderRadius: 3,
                    background: BAR_COLORS[ev.type] ?? BAR_COLORS.milestone,
                    fontSize: 9,
                    color: 'white',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontWeight: 600,
                    overflow: 'hidden',
                    whiteSpace: 'nowrap',
                    minWidth: 4,
                  }}
                >
                  {width > 10 ? ev.label : ''}
                </div>
              </div>
            </div>
          );
        })}

        <div
          style={{
            position: 'absolute',
            left: `calc(110px + (100% - 110px) * ${nowPercent / 100})`,
            top: 0,
            bottom: 0,
            width: 2,
            background: '#e05050',
            zIndex: 5,
          }}
        >
          <div
            style={{
              position: 'absolute',
              top: -14,
              transform: 'translateX(-50%)',
              fontSize: 9,
              color: '#e05050',
              fontWeight: 700,
            }}
          >
            Today
          </div>
        </div>
      </div>
    </div>
  );
}
