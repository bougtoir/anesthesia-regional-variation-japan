/**
 * Lightweight Server-Sent Events broadcaster.
 *
 * Usage:
 *   const sse = new SSEBroadcaster();
 *   // In your HTTP handler for GET /api/events:
 *   sse.addClient(res);
 *   // Push to every connected browser:
 *   sse.broadcast('project-updated', { id: '...' });
 */

export class SSEBroadcaster {
  #clients = new Set();

  /** Register a new SSE connection. */
  addClient(res) {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
      'Access-Control-Allow-Origin': '*',
    });
    res.write(':\n\n'); // comment to keep connection alive

    this.#clients.add(res);
    res.on('close', () => this.#clients.delete(res));
  }

  /** Send an event to all connected clients. */
  broadcast(event, data) {
    const payload = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
    for (const client of this.#clients) {
      client.write(payload);
    }
  }

  get clientCount() {
    return this.#clients.size;
  }
}
