export async function onRequestPost(context) {
  try {
    const { url } = await context.request.json();
    const u = new URL(url);
    if (u.protocol !== 'https:' || !u.hostname.toLowerCase().endsWith('.sharepoint.com')) {
      return new Response('Invalid SharePoint URL', { status: 400 });
    }
    u.searchParams.delete('web');
    u.searchParams.delete('action');
    u.searchParams.set('download', '1');
    const r = await fetch(u.toString(), { redirect: 'follow' });
    if (!r.ok) return new Response('SharePoint fetch failed', { status: 502 });
    const ct = (r.headers.get('content-type') || '').toLowerCase();
    if (ct.includes('text/html')) return new Response('SharePoint returned HTML', { status: 502 });
    const buf = await r.arrayBuffer();
    if (buf.byteLength < 100 || buf.byteLength > 20 * 1024 * 1024) return new Response('Invalid file size', { status: 502 });
    return new Response(buf, { headers: {
      'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'cache-control': 'no-store',
      'x-content-type-options': 'nosniff'
    }});
  } catch (e) {
    return new Response('Unable to fetch workbook', { status: 500 });
  }
}
