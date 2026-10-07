from datetime import date
from collections import Counter
import unittest

def reconcile(tickets, logs):
    answer_counts = Counter(str(log.get('answer') or '').strip() for log in logs if str(log.get('answer') or '').strip())
    logs = [log for log in logs if str(log.get('answer') or '').strip() and answer_counts[str(log.get('answer') or '').strip()] == 1]
    used_tickets, used_logs = set(), set()
    results = [None] * len(tickets)
    for same_day in (True, False):
        changed = True
        while changed:
            changed = False
            for ti, ticket in enumerate(tickets):
                if ti in used_tickets:
                    continue
                candidates = [li for li, log in enumerate(logs) if li not in used_logs and log['car'] == ticket['car'] and log['amount'] == ticket['amount'] and (log['day'] == ticket['day'] if same_day else log['day'] > ticket['day'])]
                if len(candidates) != 1:
                    continue
                li = candidates[0]
                reverse = [other for other, candidate_ticket in enumerate(tickets) if other not in used_tickets and candidate_ticket['car'] == logs[li]['car'] and candidate_ticket['amount'] == logs[li]['amount'] and (candidate_ticket['day'] == logs[li]['day'] if same_day else candidate_ticket['day'] < logs[li]['day'])]
                if len(reverse) == 1:
                    used_tickets.add(ti); used_logs.add(li)
                    results[ti] = ('同日一致' if same_day else '後日一致', logs[li]['answer'])
                    changed = True
    for ti, ticket in enumerate(tickets):
        if results[ti] is None:
            candidates = [log for li, log in enumerate(logs) if li not in used_logs and log['car'] == ticket['car'] and log['amount'] == ticket['amount'] and log['day'] >= ticket['day']]
            results[ti] = ('候補複数・要確認' if candidates else '未照合', None)
    return results

class T(unittest.TestCase):
 def test_same_day_priority_and_order_independence(self):
  tickets=[{'car':'7391','amount':25,'day':date(2026,6,10)},{'car':'7391','amount':25,'day':date(2026,6,11)}]
  logs=[{'answer':'601','car':'7391','amount':25,'day':date(2026,6,11)}]
  self.assertEqual(reconcile(tickets,logs),[('未照合',None),('同日一致','601')])
  self.assertEqual(reconcile(tickets[::-1],logs),[('同日一致','601'),('未照合',None)])
 def test_ambiguous_later_candidates_and_one_to_one(self):
  t=[{'car':'7391','amount':25,'day':date(2026,6,10)}]
  self.assertEqual(reconcile(t,[{'answer':'1','car':'7391','amount':25,'day':date(2026,6,11)},{'answer':'2','car':'7391','amount':25,'day':date(2026,6,12)}])[0][0],'候補複数・要確認')
  self.assertEqual(reconcile(t*2,[{'answer':'1','car':'7391','amount':25,'day':date(2026,6,12)}])[0][0],'候補複数・要確認')
 def test_vehicle_match_is_generic_and_no_reverse_leak_check(self):
  self.assertEqual(reconcile([{'car':'7391','amount':25,'day':date(2026,6,10)}],[{'answer':'1','car':'1889','amount':25,'day':date(2026,6,10)}])[0][0],'未照合')
  self.assertEqual(reconcile([{'car':'1889','amount':30,'day':date(2026,6,1)}],[{'answer':'2','car':'1889','amount':30,'day':date(2026,6,1)}]),[('同日一致','2')])
 def test_same_day_ambiguity_and_cross_vehicle_isolation(self):
  t=[{'car':'7391','amount':25,'day':date(2026,6,10)}]
  duplicate_day=[{'answer':'1','car':'7391','amount':25,'day':date(2026,6,10)},{'answer':'2','car':'7391','amount':25,'day':date(2026,6,10)}]
  self.assertEqual(reconcile(t,duplicate_day)[0][0],'候補複数・要確認')
  tickets=[{'car':'7391','amount':25,'day':date(2026,6,10)},{'car':'1889','amount':25,'day':date(2026,6,11)}]
  logs=[{'answer':'a','car':'1889','amount':25,'day':date(2026,6,10)},{'answer':'b','car':'7391','amount':25,'day':date(2026,6,11)}]
  self.assertEqual(reconcile(tickets,logs),[('後日一致','b'),('未照合',None)])
 def test_blank_or_duplicate_log_answers_are_not_matchable(self):
  ticket=[{'car':'7391','amount':25,'day':date(2026,6,10)}]
  self.assertEqual(reconcile(ticket,[{'answer':'','car':'7391','amount':25,'day':date(2026,6,10)}])[0][0],'未照合')
  duplicates=[{'answer':'601','car':'7391','amount':25,'day':date(2026,6,10)},{'answer':'601','car':'7391','amount':25,'day':date(2026,6,10)}]
  self.assertEqual(reconcile(ticket,duplicates)[0][0],'未照合')
 def test_vba_rejects_duplicate_and_blank_answers_and_preserves_input_columns(self):
  from pathlib import Path
  source=(Path(__file__).parents[1] / 'FuelReconciliation.bas').read_text(encoding='utf-8')
  self.assertIn('If Len(answer) > 0 Then',source)
  self.assertIn('If CLng(answerCounts(answer)) = 1 Then',source)
  self.assertIn('lastRow = LastTicketInputRow(ws)',source)
  self.assertNotRegex(source,r'(?im)^\s*ws\.Cells\([^\n]*,\s*"[ABC]"\)\s*=')
  self.assertIn('If activeRow And hasOutput And Trim$(CStr(ws.Cells(r, "H").Value)) <> OUTPUT_MARKER Then',source)
 def test_current_vba_requires_replacement(self):
  from pathlib import Path
  source=(Path(__file__).parents[1] / 'FuelReconciliation.bas').read_text(encoding='utf-8')
  self.assertNotIn('TARGET_CAR',source)
  self.assertNotIn('TARGET_START',source)
  self.assertNotIn('GetResultSheet',source)
  self.assertIn('ResolveStage',source)
if __name__=='__main__': unittest.main()
