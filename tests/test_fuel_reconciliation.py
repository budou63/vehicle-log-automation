from datetime import date
from collections import Counter
import unittest

def reconcile(tickets, logs):
    answer_counts = Counter(str(log.get('answer') or '').strip() for log in logs if str(log.get('answer') or '').strip())
    logs = [log for log in logs if str(log.get('answer') or '').strip() and answer_counts[str(log.get('answer') or '').strip()] == 1]
    used_tickets, used_logs = set(), set()
    reserved_tickets, reserved_logs = set(), set()
    results = [None] * len(tickets)
    for same_day in (True, False):
        changed = True
        while changed:
            changed = False
            for ti, ticket in enumerate(tickets):
                if ti in used_tickets:
                    continue
                candidates = [li for li, log in enumerate(logs) if li not in used_logs and (same_day or li not in reserved_logs) and log['car'] == ticket['car'] and log['amount'] == ticket['amount'] and (log['day'] == ticket['day'] if same_day else log['day'] > ticket['day'])]
                if len(candidates) != 1:
                    continue
                li = candidates[0]
                reverse = [other for other, candidate_ticket in enumerate(tickets) if other not in used_tickets and (same_day or other not in reserved_tickets) and candidate_ticket['car'] == logs[li]['car'] and candidate_ticket['amount'] == logs[li]['amount'] and (candidate_ticket['day'] == logs[li]['day'] if same_day else candidate_ticket['day'] < logs[li]['day'])]
                if len(reverse) == 1:
                    used_tickets.add(ti); used_logs.add(li)
                    results[ti] = ('同日一致' if same_day else '後日一致', logs[li]['answer'])
                    changed = True
        if same_day:
            for ti, ticket in enumerate(tickets):
                if ti not in used_tickets:
                    exact_logs = [li for li, log in enumerate(logs) if li not in used_logs and log['car'] == ticket['car'] and log['amount'] == ticket['amount'] and log['day'] == ticket['day']]
                    if exact_logs:
                        reserved_tickets.add(ti)
                        reserved_logs.update(exact_logs)
    for ti, ticket in enumerate(tickets):
        if results[ti] is None:
            if ti in reserved_tickets:
                results[ti] = ('候補複数・要確認', None)
                continue
            candidates = [log for li, log in enumerate(logs) if li not in used_logs and li not in reserved_logs and log['car'] == ticket['car'] and log['amount'] == ticket['amount'] and log['day'] >= ticket['day']]
            results[ti] = ('候補複数・要確認' if candidates else '未照合', None)
    return results

class T(unittest.TestCase):
 def test_ambiguous_same_day_records_are_reserved_from_later_matching(self):
  tickets=[{'car':'7391','amount':25,'day':date(2026,6,10)},{'car':'7391','amount':25,'day':date(2026,6,11)},{'car':'7391','amount':25,'day':date(2026,6,11)}]
  logs=[{'answer':'601','car':'7391','amount':25,'day':date(2026,6,11)}]
  expected=[('未照合',None),('候補複数・要確認',None),('候補複数・要確認',None)]
  self.assertEqual(reconcile(tickets,logs),expected)
  permutation=[tickets[2],tickets[0],tickets[1]]
  self.assertEqual(reconcile(permutation,logs),[expected[2],expected[0],expected[1]])
 def test_ordinary_unique_later_match_is_preserved(self):
  tickets=[{'car':'7391','amount':25,'day':date(2026,6,10)}]
  logs=[{'answer':'601','car':'7391','amount':25,'day':date(2026,6,11)}]
  self.assertEqual(reconcile(tickets,logs),[('後日一致','601')])
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
  source=(Path(__file__).parents[1] / 'Module1').read_text(encoding='utf-8')
  self.assertIn('If Len(answer) > 0 Then',source)
  self.assertIn('If CLng(answerCounts(answer)) = 1 Then',source)
  self.assertIn('lastRow = LastTicketInputRow(ws)',source)
  first=source.index('ResolveStage tickets, logs, usedTickets, usedLogs, matches, reservedSameDayTickets, reservedSameDayLogs, True')
  reserve=source.index('ReserveAmbiguousSameDay tickets, logs, usedTickets, usedLogs, reservedSameDayTickets, reservedSameDayLogs')
  second=source.index('ResolveStage tickets, logs, usedTickets, usedLogs, matches, reservedSameDayTickets, reservedSameDayLogs, False')
  self.assertLess(first,reserve)
  self.assertLess(reserve,second)
  self.assertIn('Not reservedTickets.Exists(ticketKey)',source)
  self.assertIn('Not reservedLogs.Exists(CStr(logItem(0)))',source)
  self.assertNotRegex(source,r'(?im)^\s*ws\.Cells\([^\n]*,\s*"[ABC]"\)\s*=')
  self.assertIn('If activeRow And hasOutput And Trim$(CStr(ws.Cells(r, "H").Value)) <> FUEL_OUTPUT_MARKER Then',source)
 def test_module1_procedure_and_for_next_structure(self):
  from pathlib import Path
  import re
  source=(Path(__file__).parents[1] / 'Module1').read_text(encoding='utf-8')
  code='\n'.join(re.sub(r"'.*$",'',re.sub(r'"(?:[^"]|"")*"','""',line)) for line in source.splitlines())
  declarations=len(re.findall(r'(?im)^\s*(?:(?:Public|Private)\s+)?(?:Sub|Function)\s+\w+',code))
  endings=len(re.findall(r'(?im)^\s*End\s+(?:Sub|Function)\b',code))
  for_loops=len(re.findall(r'(?i)\bFor\s+(?:Each\s+)?\w+\s*(?:=|In)',code))
  next_loops=len(re.findall(r'(?i)\bNext\b',code))-len(re.findall(r'(?i)\bOn\s+Error\s+Resume\s+Next\b',code))
  self.assertEqual(declarations,endings)
  self.assertEqual(for_loops,next_loops)
  self.assertRegex(code,r'(?is)For Each key In statuses\.Keys\s+If Not allRows\.Exists\(CStr\(key\)\) Then\s+allRows\.Add CStr\(key\), True\s+End If\s+Next key')
  self.assertRegex(code,r'(?is)For Each key In allRows\.Keys.*?Next key')
 def test_current_vba_requires_replacement(self):
  from pathlib import Path
  source=(Path(__file__).parents[1] / 'Module1').read_text(encoding='utf-8')
  self.assertNotIn('TARGET_CAR',source)
  self.assertNotIn('TARGET_START',source)
  self.assertNotIn('GetResultSheet',source)
  self.assertIn('ResolveStage',source)
if __name__=='__main__': unittest.main()
