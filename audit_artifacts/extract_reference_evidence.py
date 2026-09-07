import requests,bs4,re
urls={
'ref9':'https://www.researchandmarkets.com/reports/5595800/femtech-market-size-share-and-trends-analysis',
'ref10':'https://www.nature.com/articles/s44222-024-00253-7',
'ref11':'https://www.nature.com/articles/s41551-026-01758-9',
'ref13':'https://www.who.int/news-room/fact-sheets/detail/endometriosis',
'ref14':'https://www.who.int/news-room/fact-sheets/detail/polycystic-ovary-syndrome',
'ref15':'https://flo.health/newsroom/flo-health-menstrual-cycle-patterns-study',
'ref16':'https://flo.health/newsroom/flo-health-raises-over-200m',
}
terms=['39.29','97.25','16.3','2024','2030','5%','2.5','797','798','17 July','1265','1266','190','10','13','70 million','380 million','200M','unicorn','July 30','June 19']
for n,u in urls.items():
 r=requests.get(u,timeout=45,headers={'User-Agent':'Mozilla/5.0'});s=bs4.BeautifulSoup(r.text,'html.parser').get_text(' ',strip=True);print('\n##',n,'status',r.status_code,'title',bs4.BeautifulSoup(r.text,'html.parser').title.get_text(' ',strip=True) if bs4.BeautifulSoup(r.text,'html.parser').title else '')
 for t in terms:
  m=re.search(r'.{0,150}'+re.escape(t)+r'.{0,250}',s,re.I)
  if m:print(t,':',m.group(0))
