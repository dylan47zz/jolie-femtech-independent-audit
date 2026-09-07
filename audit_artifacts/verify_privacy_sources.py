import requests,re,bs4
urls={
'Mozilla':'https://www.mozillafoundation.org/en/blog/in-post-roe-v-wade-era-mozilla-labels-18-of-25-popular-period-and-pregnancy-tracking-tech-with-privacy-not-included-warning/',
'FTC_Flo':'https://www.ftc.gov/news-events/news/press-releases/2021/06/ftc-finalizes-order-flo-health-fertility-tracking-app-shared-sensitive-health-data-facebook-google',
'FTC_Premom':'https://www.ftc.gov/news-events/news/press-releases/2023/05/ovulation-tracking-app-premom-will-be-barred-sharing-health-data-advertising-under-proposed-ftc',
'Flo':'https://flo.health/newsroom/flo-response-ftc-settlement-update'
}
terms=['18','25','20','five','5','warning','August','2022','Facebook','Google','independent','review','March','2022','proposed','100,000','$100,000','civil penalty','final','order','advertising']
for n,u in urls.items():
 r=requests.get(u,timeout=45,headers={'User-Agent':'Mozilla/5.0'});s=bs4.BeautifulSoup(r.text,'html.parser').get_text(' ',strip=True);print('\n###',n,r.status_code,len(s))
 for t in terms:
  m=re.search(r'.{0,210}'+re.escape(t)+r'.{0,330}',s,re.I)
  if m:print('TERM',t,':',m.group(0))
