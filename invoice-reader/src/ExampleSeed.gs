/** Sample starting examples (made up for the public demo). Used once by "Load starting examples".
 *  The real version is seeded from a few hundred invoices checked by a person. */
const SEED_EXAMPLES = [
  {"supplier":"Greenleaf Produce Ltd","invoice_no":"104233","invoice_no_note":"","doc_type":"invoice","date":"02/05/2026","due_date":"","restaurant":"Noodle Bar Westside","lines":[{"category":"Food","net":52.40,"vat_pct":0,"vat":0}],"gross":52.40,"details":"Broccoli, cabbage, carrots, peppers","notes":"Invoice no. is the 6-digit number top right, not the order number","scan":"sample p1"},
  {"supplier":"GasCo LPG Ltd","invoice_no":"585729","invoice_no_note":"","doc_type":"invoice","date":"30/04/2026","due_date":"30/04/2026","restaurant":"Noodle Bar Central","lines":[{"category":"LPG gas & charcoal","net":1122.32,"vat_pct":5,"vat":56.12}],"gross":1178.44,"details":"LPG in bulk, 586 kg (5 deliveries April)","notes":"","scan":"sample p2"},
  {"supplier":"Vineyard Wines & Spirits Ltd","invoice_no":"INV-2026-0412","invoice_no_note":"","doc_type":"credit note","date":"12/05/2026","due_date":"","restaurant":"Harbour Grill","lines":[{"category":"Drinks","net":-48.00,"vat_pct":19,"vat":-9.12}],"gross":-57.12,"details":"Return of 6 bottles house red","notes":"","scan":"sample p3"}
];
