select c.c_comp_co, s.C_ACTNO, s.C_ACTNAM, c.C_ADRMOD_BONO, to_char(c.C_MOD_DT,'yyyymmdd')
from CHINESEDATAMODIFY c, shareholder s
where c.C_COMP_CO = s.C_COMP_CO
and  c.C_ACTNO = s.C_ACTNO
and c.C_ADRMOD_BONO is not null
and c.c_comp_co='101'
and to_char(c.C_MOD_DT,'yyyymmdd') between '20250301' and '20250331'
group by c.c_comp_co, s.C_ACTNO, s.C_ACTNAM, s.C_ACTNAM, c.C_ADRMOD_BONO, to_char(c.C_MOD_DT,'yyyymmdd')
order by c_comp_co ,    c.C_ADRMOD_BONO