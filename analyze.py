#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8')

exec(open('matching.py', encoding='utf-8').read().replace('if __name__ == "__main__":', 'if False:'))

buyers, bw = load_buyers('2026 구매담당자 리스트.xlsx')
companies, cw = load_companies('2026 참가기업 리스트.xlsx', buyers)
pool1, pool2, pool3, pool4, pool5 = classify_pairs(buyers, companies)
matches_by_round, matched, match_type, slots = greedy_matching(buyers, companies, pool1, pool2, pool3, pool4, pool5)

buyer_match_count = {bid: 0 for bid in buyers}
for r in [1,2,3,4]:
    for bid, cname in matches_by_round[r]:
        buyer_match_count[bid] += 1

dist = {0:0, 1:0, 2:0, 3:0, 4:0}
for cnt in buyer_match_count.values():
    dist[cnt] += 1

print('=== 바이어별 매칭 수 분포 ===')
for k, v in dist.items():
    print(f'  {k}라운드 매칭: {v}명')

print()
print('=== 수학적 한계 분석 ===')
print(f'  기업 수: {len(companies)}개')
print(f'  최대 가능 총 매칭: {len(companies)} x 4라운드 = {len(companies)*4}쌍')
print(f'  바이어 수: {len(buyers)}명')
print(f'  바이어 전원 4라운드 충족에 필요한 슬롯: {len(buyers)*4}개')
print(f'  실제 매칭: {len(matched)}쌍')

print()
for date in ['2026-05-21', '2026-05-22']:
    date_buyers = [b for b in buyers.values() if b['date'] == date]
    date_companies = [c for c in companies.values() if date in c['dates']]
    needed = len(date_buyers) * 4
    available = len(date_companies) * 4
    print(f'{date}: 바이어 {len(date_buyers)}명 / 참가기업 {len(date_companies)}개')
    print(f'  → 필요 슬롯 {needed}개 vs 최대 가능 {available}개 (부족: {max(0, needed-available)}개)')

print()
print('=== 0라운드 매칭 바이어 ===')
for bid, cnt in buyer_match_count.items():
    if cnt == 0:
        b = buyers[bid]
        print(f'  {bid} ({b["org"][:20]} / {b["name"]}): 관심기업 {len(b["interested_companies"])}개')
