import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { createMicrosoftDraft, draftPayload, DraftOutcomeUnknown, MAX_DRAFT_ATTACHMENT } from '../src/microsoft-draft';
import type { Comparison } from '../src/types';

const rows: Comparison[] = [{ id: '1', severity: 'critical', bomPart: 'A', jobPart: 'A', bomQty: '1', jobQty: '2', bomThickness: '2', jobThickness: '2', explanation: 'Quantity differs.' }];
const context = { bomName: 'BOM.xlsx', jobName: 'Job.doc' };
test('draft payload preserves attachment bytes and exact recipients for both formats', async () => {
  for (const format of ['pdf', 'csv'] as const) {
    const report = new Blob([new Uint8Array([0, 255, 10, 13, 65])]);
    const payload = await draftPayload(report, format, rows, context);
    assert.deepEqual(payload.toRecipients.map(r => r.emailAddress.address), ['werner@imssa.co.za', 'john@imssa.co.za']);
    assert.equal(payload.ccRecipients[0].emailAddress.address, 'nikita@imssa.co.za');
    assert.deepEqual(Buffer.from(payload.attachments[0].contentBytes, 'base64'), Buffer.from(await report.arrayBuffer()));
    assert.equal(payload.attachments[0].name, `ims-sheet-metal-audit.${format}`);
    assert.match(payload.body.content, /Selected rows: 1 \(critical: 1/);
  }
});
test('creates one draft request including attachment, never calls send', async () => {
  let calls = 0;
  const request: typeof fetch = async (url, init) => {
    calls++;
    assert.equal(url, 'https://graph.microsoft.com/v1.0/me/messages');
    assert.equal(init?.method, 'POST');
    assert.equal(new Headers(init?.headers).get('Authorization'), 'Bearer test-token');
    assert.equal(JSON.parse(String(init?.body)).attachments.length, 1);
    return new Response(JSON.stringify({ id: 'draft-id', isDraft: true }), { status: 201 });
  };
  assert.equal(await createMicrosoftDraft('test-token', new Blob(['report']), 'csv', rows, context, request), 'draft-id');
  assert.equal(calls, 1);
});
test('network uncertainty never auto-retries creation', async () => {
  let calls = 0;
  await assert.rejects(createMicrosoftDraft('test', new Blob(['report']), 'csv', rows, context, async () => { calls++; throw new Error('network'); }), DraftOutcomeUnknown);
  assert.equal(calls, 1);
});
test('permission errors are actionable and oversized attachments make no request', async () => {
  await assert.rejects(createMicrosoftDraft('test', new Blob(['r']), 'csv', rows, context, async () => new Response('', { status: 403 })), /Mail.ReadWrite/);
  await assert.rejects(createMicrosoftDraft('test', new Blob([new Uint8Array(MAX_DRAFT_ATTACHMENT + 1)]), 'pdf', rows, context, async () => { assert.fail('must not upload oversized report'); }), /too large/);
});
test('malformed successful responses are treated as uncertain', async () => {
  await assert.rejects(createMicrosoftDraft('test', new Blob(['r']), 'csv', rows, context, async () => new Response('invalid json', { status: 201 })), DraftOutcomeUnknown);
});
