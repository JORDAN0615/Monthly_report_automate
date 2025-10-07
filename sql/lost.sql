select c_comp_co,c_ap_actno,c_tra_dt,C_TRA_BONO from LOSSTRANSACTIONMASTER 
where c_ty='0' 
and c_comp_co='101' 
and to_char(c_tra_dt,'yyyymmdd') between '20250301' and '20250331'
order by c_comp_co,c_ap_actno