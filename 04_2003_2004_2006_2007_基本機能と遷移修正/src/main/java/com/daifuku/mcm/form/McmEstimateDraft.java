package com.daifuku.mcm.form;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.*;

/** 2002で選定した見積一式。登録まではセッションだけに保持する。 */
public class McmEstimateDraft implements Serializable {
    private static final long serialVersionUID = 1L;
    private final Map<String,List<Map<String,Object>>> tables = new LinkedHashMap<>();
    private Mcm2003uForm form;
    private BigDecimal originalId;
    private String originalVersion;
    public Map<String,List<Map<String,Object>>> getTables(){return tables;}
    public List<Map<String,Object>> rows(String table){return tables.computeIfAbsent(table,k->new ArrayList<>());}
    public Mcm2003uForm getForm(){return form;}
    public void setForm(Mcm2003uForm v){form=v;}
    public BigDecimal getOriginalId(){return originalId;}
    public void setOriginalId(BigDecimal v){originalId=v;}
    public String getOriginalVersion(){return originalVersion;}
    public void setOriginalVersion(String v){originalVersion=v;}
}
