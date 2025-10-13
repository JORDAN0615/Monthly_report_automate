select c_comp_co,c_ap_actno,c_tra_dt,C_TRA_BONO from LOSSTRANSACTIONMASTER
where c_ty='0'
and c_comp_co = :company_code
and to_char(c_tra_dt,'yyyymmdd') between :start_date and :end_date
order by c_comp_co,c_ap_actno