import { useEffect, useRef } from 'react';

type EventHandler = (data: unknown) => void;

/**
 * Hook that connects to the server's SSE endpoint for real-time updates.
 * Falls back gracefully when no server is running (pure static mode).
 *
 * @param handlers Map of event names → handler functions
 * @param enabled  Whether to connect (default: true)
 */
export function useServerEvents(
  handlers: Record<string, EventHandler>,
  enabled = true,
) {
  const handlersRef = useRef(handlers);
  handlersRef.current = handlers;

  useEffect(() => {
    if (!enabled) return;

    const sseUrl = `${window.location.origin}/api/events`;
    let es: EventSource | null = null;
    let retryTimer: ReturnType<typeof setTimeout>;

    function connect() {
      try {
        es = new EventSource(sseUrl);

        es.onopen = () => {
          console.log('[SSE] connected');
        };

        es.onerror = () => {
          // Silently reconnect — this is expected when running in static mode
          es?.close();
          retryTimer = setTimeout(connect, 5000);
        };

        // Register all handlers
        for (const event of Object.keys(handlersRef.current)) {
          es.addEventListener(event, (e: MessageEvent) => {
            try {
              const data = JSON.parse(e.data);
              handlersRef.current[event]?.(data);
            } catch {
              // ignore parse errors
            }
          });
        }
      } catch {
        // EventSource not available or network error — retry later
        retryTimer = setTimeout(connect, 5000);
      }
    }

    connect();

    return () => {
      es?.close();
      clearTimeout(retryTimer);
    };
  }, [enabled]);
}
