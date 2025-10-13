select c.c_comp_co, s.C_ACTNO, s.C_ACTNAM, c.C_ADRMOD_BONO, to_char(c.C_MOD_DT,'yyyymmdd')
from CHINESEDATAMODIFY c, shareholder s
where c.C_COMP_CO = s.C_COMP_CO
and  c.C_ACTNO = s.C_ACTNO
and c.C_ADRMOD_BONO is not null
and c.c_comp_co = :company_code
and to_char(c.C_MOD_DT,'yyyymmdd') between :start_date and :end_date
group by c.c_comp_co, s.C_ACTNO, s.C_ACTNAM, s.C_ACTNAM, c.C_ADRMOD_BONO, to_char(c.C_MOD_DT,'yyyymmdd')
order by c_comp_co ,    c.C_ADRMOD_BONO