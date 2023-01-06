```python
import pandas as pd
from pathlib import Path

io = r'..\Level 3 &VS scorecard.xlsx'
```


```python
data = pd.read_excel(io, sheet_name = 'L3&VS-Assy',header=2)
data.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>#</th>
      <th>区域</th>
      <th>PQVC</th>
      <th>指标分类1</th>
      <th>指标名称</th>
      <th>Data provider</th>
      <th>Unnamed: 6</th>
      <th>2020 Baseline</th>
      <th>YTD</th>
      <th>Jan</th>
      <th>...</th>
      <th>Unnamed: 24</th>
      <th>Unnamed: 25</th>
      <th>Unnamed: 26</th>
      <th>Unnamed: 27</th>
      <th>Unnamed: 28</th>
      <th>Unnamed: 29</th>
      <th>Unnamed: 30</th>
      <th>Unnamed: 31</th>
      <th>Unnamed: 32</th>
      <th>Unnamed: 33</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1</td>
      <td>Assy</td>
      <td>P</td>
      <td>突破性指标</td>
      <td>RIF</td>
      <td>Jerry Sun</td>
      <td>Plan</td>
      <td>NaN</td>
      <td>0</td>
      <td>0</td>
      <td>...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>1</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>Actual</td>
      <td>0.29</td>
      <td>0</td>
      <td>0</td>
      <td>...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>2</th>
      <td>2</td>
      <td>Assy</td>
      <td>P</td>
      <td>突破性指标</td>
      <td>LTCFR</td>
      <td>NaN</td>
      <td>Plan</td>
      <td>NaN</td>
      <td>0</td>
      <td>0</td>
      <td>...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>3</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>Actual</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>4</th>
      <td>3</td>
      <td>Assy</td>
      <td>P</td>
      <td>基础性指标</td>
      <td>EI Engagement Index (%)</td>
      <td>NaN</td>
      <td>Plan</td>
      <td>NaN</td>
      <td>0.9</td>
      <td>NaN</td>
      <td>...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
<p>5 rows × 34 columns</p>
</div>




```python

```


```python
data = pd.read_excel(io, sheet_name = 'L3&VS-Assy',header=2,usecols=['指标分类1','指标名称', 'Jan', 'Feb', 'Mar', 'Apr','May','Jun','Jul','Aug','Sep','Oct'])
data.head()
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>指标分类1</th>
      <th>指标名称</th>
      <th>Jan</th>
      <th>Feb</th>
      <th>Mar</th>
      <th>Apr</th>
      <th>May</th>
      <th>Jun</th>
      <th>Jul</th>
      <th>Aug</th>
      <th>Sep</th>
      <th>Oct</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>突破性指标</td>
      <td>RIF</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>1</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>2</th>
      <td>突破性指标</td>
      <td>LTCFR</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>3</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>4</th>
      <td>基础性指标</td>
      <td>EI Engagement Index (%)</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
</div>




```python
data_1=data.copy()
data_1['指标名称']=data_1['指标名称'].fillna(method="ffill",limit=1)
print(data_1)
```

         指标分类1                                   指标名称       Jan       Feb  \
    0    突破性指标                                    RIF         0         0   
    1      NaN                                    RIF         0         0   
    2    突破性指标                                  LTCFR         0         0   
    3      NaN                                  LTCFR         0         0   
    4    基础性指标                EI Engagement Index (%)       NaN       NaN   
    ..     ...                                    ...       ...       ...   
    519    NaN  Total Variable Cost ($ rate per hour)       NaN       NaN   
    520  基础性指标                             OPACC ($M)   0.0003    0.0082    
    521    NaN                             OPACC ($M)   0.0003    0.0082    
    522    NaN                    Achieve NPI  target       NaN       NaN   
    523    NaN                    Achieve NPI  target       NaN       NaN   
    
              Mar       Apr       May       Jun       Jul       Aug       Sep  \
    0           0         0         0         0         0         0         0   
    1           0         0         0         0         0         0         0   
    2           0         0         0         0         0         0         0   
    3           0         0         0         0         0         0         0   
    4         NaN       NaN       NaN       NaN       NaN       NaN       NaN   
    ..        ...       ...       ...       ...       ...       ...       ...   
    519       NaN       NaN       NaN       NaN       NaN       NaN       NaN   
    520   0.0218    0.0371    0.0525    0.0678    0.1567    0.2363    0.3507    
    521   0.0218    0.0371    0.0525        NaN       NaN       NaN       NaN   
    522       NaN       NaN       NaN       NaN       NaN       NaN       NaN   
    523       NaN       NaN       NaN       NaN       NaN       NaN       NaN   
    
              Oct  
    0           0  
    1           0  
    2           0  
    3           0  
    4         NaN  
    ..        ...  
    519       NaN  
    520   0.4200   
    521       NaN  
    522       NaN  
    523       NaN  
    
    [524 rows x 12 columns]
    


```python
sheet_index1=data[data['指标分类1'].isin(['LHEX- SUB1'])].index.tolist()
print(sheet_index1)
sheet_index2=data[data['指标分类1'].isin(['LHEX- SUB2'])].index.tolist()
print(sheet_index2)
if len(sheet_index1)!=1:
    print('index error')
```

    [64]
    [130]
    


```python
sheet1=data[:][sheet_index1[0]:sheet_index2[0]]
sheet2=data_1[:][sheet_index1[0]:sheet_index2[0]]
sheet2
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>指标分类1</th>
      <th>指标名称</th>
      <th>Jan</th>
      <th>Feb</th>
      <th>Mar</th>
      <th>Apr</th>
      <th>May</th>
      <th>Jun</th>
      <th>Jul</th>
      <th>Aug</th>
      <th>Sep</th>
      <th>Oct</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>64</th>
      <td>LHEX- SUB1</td>
      <td>李保平</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>65</th>
      <td>指标分类1</td>
      <td>指标名称</td>
      <td>Jan</td>
      <td>Feb</td>
      <td>Mar</td>
      <td>Apr</td>
      <td>May</td>
      <td>Jun</td>
      <td>Jul</td>
      <td>Aug</td>
      <td>Sep</td>
      <td>Oct</td>
    </tr>
    <tr>
      <th>66</th>
      <td>突破性指标</td>
      <td>RIF(可记录伤害）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>67</th>
      <td>NaN</td>
      <td>RIF(可记录伤害）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>68</th>
      <td>突破性指标</td>
      <td>LTCFR（损失工作日）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>...</th>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
      <td>...</td>
    </tr>
    <tr>
      <th>125</th>
      <td>NaN</td>
      <td>OPACC ($M)</td>
      <td>0.0007</td>
      <td>0.0014</td>
      <td>0.0030</td>
      <td>0.0046</td>
      <td>0.0062</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>126</th>
      <td>NaN</td>
      <td>BVVGB BGB BGBGBB VBB FVBVBB VBBB VGVVB B VVFF...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>127</th>
      <td>NaN</td>
      <td>BVVGB BGB BGBGBB VBB FVBVBB VBBB VGVVB B VVFF...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>128</th>
      <td>NaN</td>
      <td>BVVGB BGB BGBGBB VBB FVBVBB VBBB VGVVB B VVFF...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>129</th>
      <td>NaN</td>
      <td>BVVGB BGB BGBGBB VBB FVBVBB VBBB VGVVB B VVFF...</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
<p>66 rows × 12 columns</p>
</div>




```python
series_1=sheet1['指标名称'].isin(['RIF(可记录伤害）','LTCFR（损失工作日）','EI Engagement Index (%)','ABBS完成率','安全不符合项及时关闭率','Missed defect (YE)- Assy','Efficiency improvement-Assy','# of operators OT > 36H/M'])
print(series_1)
```

    64     False
    65     False
    66      True
    67     False
    68      True
           ...  
    125    False
    126    False
    127    False
    128    False
    129    False
    Name: 指标名称, Length: 66, dtype: bool
    


```python
series_2=series_1.copy()
print(type(series_1))
for key,value in enumerate(series_1,start=sheet_index1[0]):
    if value==True:
        series_2[key]=False
        series_2[key+1]=True
#     print(key,value)
print(series_2)
    
```

    <class 'pandas.core.series.Series'>
    64     False
    65     False
    66     False
    67      True
    68     False
           ...  
    125    False
    126    False
    127    False
    128    False
    129    False
    Name: 指标名称, Length: 66, dtype: bool
    


```python
sheet1.loc[series_1,:]
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>指标分类1</th>
      <th>指标名称</th>
      <th>Jan</th>
      <th>Feb</th>
      <th>Mar</th>
      <th>Apr</th>
      <th>May</th>
      <th>Jun</th>
      <th>Jul</th>
      <th>Aug</th>
      <th>Sep</th>
      <th>Oct</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>66</th>
      <td>突破性指标</td>
      <td>RIF(可记录伤害）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>68</th>
      <td>突破性指标</td>
      <td>LTCFR（损失工作日）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>70</th>
      <td>基础性指标</td>
      <td>EI Engagement Index (%)</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
      <td>0.9</td>
    </tr>
    <tr>
      <th>76</th>
      <td>基础性指标</td>
      <td>ABBS完成率</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
    </tr>
    <tr>
      <th>78</th>
      <td>基础性指标</td>
      <td>安全不符合项及时关闭率</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
      <td>0.95</td>
    </tr>
    <tr>
      <th>94</th>
      <td>突破性指标</td>
      <td>Missed defect (YE)- Assy</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
    </tr>
    <tr>
      <th>112</th>
      <td>突破性指标</td>
      <td>Efficiency improvement-Assy</td>
      <td>0</td>
      <td>0</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.04</td>
      <td>0.04</td>
      <td>0.04</td>
      <td>0.06</td>
      <td>0.08</td>
    </tr>
    <tr>
      <th>120</th>
      <td>突破性指标</td>
      <td># of operators OT &gt; 36H/M</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
  </tbody>
</table>
</div>




```python
sheet3=sheet2.loc[series_2,:]
sheet3
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>指标分类1</th>
      <th>指标名称</th>
      <th>Jan</th>
      <th>Feb</th>
      <th>Mar</th>
      <th>Apr</th>
      <th>May</th>
      <th>Jun</th>
      <th>Jul</th>
      <th>Aug</th>
      <th>Sep</th>
      <th>Oct</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>67</th>
      <td>NaN</td>
      <td>RIF(可记录伤害）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>69</th>
      <td>NaN</td>
      <td>LTCFR（损失工作日）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>71</th>
      <td>NaN</td>
      <td>EI Engagement Index (%)</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>77</th>
      <td>NaN</td>
      <td>ABBS完成率</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
    </tr>
    <tr>
      <th>79</th>
      <td>NaN</td>
      <td>安全不符合项及时关闭率</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
    </tr>
    <tr>
      <th>95</th>
      <td>NaN</td>
      <td>Missed defect (YE)- Assy</td>
      <td>0</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>4</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>2</td>
    </tr>
    <tr>
      <th>113</th>
      <td>NaN</td>
      <td>Efficiency improvement-Assy</td>
      <td>0</td>
      <td>0</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>121</th>
      <td>NaN</td>
      <td># of operators OT &gt; 36H/M</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>0</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
</div>




```python
sheet3.values
```




    array([[nan, 'RIF(可记录伤害）', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
           [nan, 'LTCFR（损失工作日）', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
           [nan, 'EI Engagement Index (%)', nan, nan, nan, nan, nan, nan,
            nan, nan, nan, nan],
           [nan, 'ABBS完成率', 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
           [nan, '安全不符合项及时关闭率', 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
           [nan, 'Missed defect (YE)- Assy', 0, 1, 1, 1, 4, 1, 1, 1, 1, 2],
           [nan, 'Efficiency improvement-Assy', 0, 0, 0.02, 0.02, 0.02, 0.02,
            0.02, nan, nan, nan],
           [nan, '# of operators OT > 36H/M', nan, nan, nan, nan, '0', nan,
            nan, nan, nan, nan]], dtype=object)




```python
del sheet3['指标分类1']
```


    ---------------------------------------------------------------------------

    KeyError                                  Traceback (most recent call last)

    D:\HT_APPDATA\anaconda\lib\site-packages\pandas\core\indexes\base.py in get_loc(self, key, method, tolerance)
       3079             try:
    -> 3080                 return self._engine.get_loc(casted_key)
       3081             except KeyError as err:
    

    pandas\_libs\index.pyx in pandas._libs.index.IndexEngine.get_loc()
    

    pandas\_libs\index.pyx in pandas._libs.index.IndexEngine.get_loc()
    

    pandas\_libs\hashtable_class_helper.pxi in pandas._libs.hashtable.PyObjectHashTable.get_item()
    

    pandas\_libs\hashtable_class_helper.pxi in pandas._libs.hashtable.PyObjectHashTable.get_item()
    

    KeyError: '指标分类1'

    
    The above exception was the direct cause of the following exception:
    

    KeyError                                  Traceback (most recent call last)

    <ipython-input-52-c8a80de1c4ed> in <module>
    ----> 1 del sheet3['指标分类1']
          2 sheet3
    

    D:\HT_APPDATA\anaconda\lib\site-packages\pandas\core\generic.py in __delitem__(self, key)
       3964             # there was no match, this call should raise the appropriate
       3965             # exception:
    -> 3966             loc = self.axes[-1].get_loc(key)
       3967             self._mgr.idelete(loc)
       3968 
    

    D:\HT_APPDATA\anaconda\lib\site-packages\pandas\core\indexes\base.py in get_loc(self, key, method, tolerance)
       3080                 return self._engine.get_loc(casted_key)
       3081             except KeyError as err:
    -> 3082                 raise KeyError(key) from err
       3083 
       3084         if tolerance is not None:
    

    KeyError: '指标分类1'



```python
sheet3
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>指标名称</th>
      <th>Jan</th>
      <th>Feb</th>
      <th>Mar</th>
      <th>Apr</th>
      <th>May</th>
      <th>Jun</th>
      <th>Jul</th>
      <th>Aug</th>
      <th>Sep</th>
      <th>Oct</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>67</th>
      <td>RIF(可记录伤害）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>69</th>
      <td>LTCFR（损失工作日）</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
      <td>0</td>
    </tr>
    <tr>
      <th>71</th>
      <td>EI Engagement Index (%)</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>77</th>
      <td>ABBS完成率</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
    </tr>
    <tr>
      <th>79</th>
      <td>安全不符合项及时关闭率</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
    </tr>
    <tr>
      <th>95</th>
      <td>Missed defect (YE)- Assy</td>
      <td>0</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>4</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>1</td>
      <td>2</td>
    </tr>
    <tr>
      <th>113</th>
      <td>Efficiency improvement-Assy</td>
      <td>0</td>
      <td>0</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>0.02</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>121</th>
      <td># of operators OT &gt; 36H/M</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>0</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
  </tbody>
</table>
</div>




```python
sheet3.values
```




    array([['RIF(可记录伤害）', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
           ['LTCFR（损失工作日）', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
           ['EI Engagement Index (%)', nan, nan, nan, nan, nan, nan, nan,
            nan, nan, nan],
           ['ABBS完成率', 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
           ['安全不符合项及时关闭率', 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
           ['Missed defect (YE)- Assy', 0, 1, 1, 1, 4, 1, 1, 1, 1, 2],
           ['Efficiency improvement-Assy', 0, 0, 0.02, 0.02, 0.02, 0.02,
            0.02, nan, nan, nan],
           ['# of operators OT > 36H/M', nan, nan, nan, nan, '0', nan, nan,
            nan, nan, nan]], dtype=object)




```python
for list in sheet3.values:
    if list[0]=='RIF(可记录伤害）':
        print('ok')
    elif list[0]=='RIF(可记录伤害）':
        
#     print(key,value)
# print(series_2)
# for i in enusheet3['Jan']
```

    ok
    


```python

```


```python

```


```python

```


```python

```
