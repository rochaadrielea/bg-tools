-- 002 · the default chain.
--
-- These are ROWS, not code. Add a checkpoint with an INSERT, retire one with
-- active=false, reorder by changing `position`. Nothing here needs a rebuild.
--
-- Scan and Iris sit AFTER the work order on purpose: nothing is scanned on
-- arrival — receiving is confirmed on screen against the system.

INSERT INTO station (key, name, source, icon, position) VALUES
  ('stock',    'Stock',      'SAP',               '▦',  1),
  ('reserved', 'Logistics',  'RESERVATION',       '▤',  2),
  ('shipped',  'Transport',  'DELIVERY NOTE',     '⇉',  3),
  ('received', 'Receiving',  'CONFIRMED ON SCREEN','⇩', 4),
  ('staged',   'Work order', 'STAGED',            '⚙',  5),
  ('iris',     'Iris',       'PHOTOS',            '◉',  6),
  ('scanned',  'Scan',       'LABELS',            '▥',  7),
  ('consumed', 'Consume',    'MATERIAL OUT',      '▼',  8),
  ('woclose',  'WO close',   'ALL CONSUMED',      '⊘',  9),
  ('asbuilt',  'As-built',   'ABCL',              '✓', 10)
ON CONFLICT (key) DO NOTHING;

-- Stations 1-6 answer to the MBOM: that is the list logistics and production
-- work from. The EBOM is the customer-facing check, once, at the end.
INSERT INTO gate (key, expected_ref, present_ref, label_a, label_b, position) VALUES
  ('mbom_reserved',    'bom:mbom',        'station:reserved', 'MBOM',      'Reserved',     1),
  ('reserved_shipped', 'station:reserved','station:shipped',  'Reserved',  'Shipped',      2),
  ('shipped_received', 'station:shipped', 'station:received', 'Shipped',   'Received',     3),
  ('received_staged',  'station:received','station:staged',   'Received',  'Staged',       4),
  ('staged_iris',      'station:staged',  'station:iris',     'Staged',    'Iris read',    5),
  ('iris_scanned',     'station:iris',    'station:scanned',  'Iris read', 'Scanned',      6),
  ('scanned_consumed', 'station:scanned', 'station:consumed', 'Scanned',   'Consumed',     7),
  ('consumed_woclose', 'station:consumed','station:woclose',  'Consumed',  'WO movements', 8),
  ('mbom_asbuilt',     'bom:mbom',        'station:asbuilt',  'MBOM',      'As-built',     9),
  ('ebom_asbuilt',     'bom:ebom',        'station:asbuilt',  'EBOM',      'As-built',    10)
ON CONFLICT (key) DO NOTHING;