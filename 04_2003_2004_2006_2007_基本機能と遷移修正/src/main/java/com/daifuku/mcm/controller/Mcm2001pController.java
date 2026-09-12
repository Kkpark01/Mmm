package com.daifuku.mcm.controller;
import java.io.Serializable;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import jakarta.servlet.http.HttpSession;
import org.springframework.http.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import com.daifuku.mcm.common.BaseController;
@Controller @RequestMapping("/mcm2001p")
public class Mcm2001pController extends BaseController {
    public record Report(BigDecimal estimateId,String user,byte[] bytes) implements Serializable { }
    @GetMapping("/download") public ResponseEntity<byte[]> download(HttpSession session){
        if(!(hasUpdateAuthority()||hasInspectionAuthority()))return ResponseEntity.status(403).build();
        if(!(session.getAttribute("mcm2001p.report") instanceof Report r)||!java.util.Objects.equals(r.user(),getLoginUserId()))return ResponseEntity.status(400).build();
        return ResponseEntity.ok().contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).cacheControl(CacheControl.noStore()).header(HttpHeaders.CONTENT_DISPOSITION,ContentDisposition.attachment().filename("店舗見積書類.xlsx",StandardCharsets.UTF_8).build().toString()).body(r.bytes());
    }
    protected String getScreenTitle(){return "店舗見積書類";}protected String getFunctionId(){return "MCM2001P";}
}
