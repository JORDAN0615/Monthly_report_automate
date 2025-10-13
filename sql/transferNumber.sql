select t.c_comp_co, t.C_TSF_BONO , to_char(t.C_TSF_DT,'yyyymmdd'), t.C_TRA_ACTNO ,
t.C_RECI_ACTNO , t.C_TSF_SHA
from TRANSFERTRANSACTION t
where to_char(t.C_TSF_DT,'yyyymmdd') between :start_date and :end_date
and t.c_comp_co = :company_code
and t.C_TSF_BONO is not null
order by t.c_comp_co, t.C_TSF_BONO