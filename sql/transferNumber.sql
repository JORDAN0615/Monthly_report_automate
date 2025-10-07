select t.c_comp_co, t.C_TSF_BONO , to_char(t.C_TSF_DT,'yyyymmdd'), t.C_TRA_ACTNO ,
t.C_RECI_ACTNO , t.C_TSF_SHA
from TRANSFERTRANSACTION t
where to_char(t.C_TSF_DT,'yyyymmdd') between '20250301' and '20250331'
and t.c_comp_co='101'
and t.C_TSF_BONO is not null
order by t.c_comp_co, t.C_TSF_BONO