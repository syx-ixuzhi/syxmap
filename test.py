import streamlit as st
import folium
from MyFunc import init_st
from streamlit_folium import st_folium
from MyFunc import read_sql

init_st(title='地图',icon='',layout='wide')
tiles1 = 'http://map.geoq.cn/ArcGIS/rest/services/ChinaOnlineCommunity/MapServer/tile/{z}/{y}/{x}' #智图
tiles2 = 'https://webrd02.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=7&x={x}&y={y}&z={z}' #高德

m = folium.Map(location=[31.952768,118.815159],
               tiles=tiles2,
               attr='高德地图',
               zoom_start=12,
               control_scale=True
               )
df = read_sql(f"SELECT name,address,coordinate FROM store_info WHERE class='门店' AND POSITION('停用' IN name) = 0 AND POSITION('新开门' IN name) = 0 AND POSITION('新开店' IN name) = 0;")
for store_name,address,coordinate in zip(df['name'],df['address'],df['coordinate']):
    lat_lng = coordinate.split(',')
    lat = lat_lng[1]
    lng = lat_lng[0]
    folium.Marker([lat, lng],
                  popup=folium.Popup(html='<span style="font-size:10pt;font-weight:bold">' + store_name + '</span><div style="background:lightgray;width:100%;height:1pt"></div>地址：' + address, max_width=400, show=False, sticky=False),
                  tooltip=store_name,
                  icon=folium.Icon(color='blue', prefix='fa', icon='shop')).add_to(m)

st_map = st_folium(m,width='100%',height=800)