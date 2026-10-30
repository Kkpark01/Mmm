package com.daifuku.mcm.service;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.sql.ResultSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.springframework.test.util.ReflectionTestUtils;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;
import org.springframework.jdbc.core.RowCallbackHandler;

import com.daifuku.mcm.common.CustomerScreenSupport;
import com.daifuku.mcm.common.Mcm2006uInputCheck;
import com.daifuku.mcm.form.Mcm2006uForm;
import com.daifuku.mcm.repository.Mcm2006uRepository;

class Mcm2006uPeriodLockTest {
    private static final BigDecimal CONTRACT = BigDecimal.valueOf(42);
    private static final Timestamp TIME = Timestamp.valueOf("2026-10-29 12:00:00");

    private Mcm2006uService service(Mcm2006uRepository repo) {
        var service = spy(new Mcm2006uService());
        ReflectionTestUtils.setField(service, "repo", repo);
        ReflectionTestUtils.setField(service, "integrity", mock(McmCustomerIntegrityService.class));
        doNothing().when(service).checkForSave(any());
        return service;
    }
    private Mcm2006uForm form(Map<String, Mcm2006uForm.PeriodLock> original) {
        var form = new Mcm2006uForm();
        form.setUkKeiyakuId(CONTRACT);
        form.setVersion("header-version");
        form.setOriginalPeriodLocks(original);
        return form;
    }
    private Mcm2006uRepository repo(Map<String, Mcm2006uForm.PeriodLock> current) {
        var repo = mock(Mcm2006uRepository.class);
        var header = new Mcm2006uForm();
        header.setUkKeiyakuId(CONTRACT);
        header.setVersion("header-version");
        when(repo.findKeiyaku(CONTRACT)).thenReturn(header);
        when(repo.lockPeriodVersions(CONTRACT)).thenReturn(current);
        return repo;
    }
    private Map<String, Mcm2006uForm.PeriodLock> snapshot(String support, Timestamp timestamp) {
        var map = new LinkedHashMap<String, Mcm2006uForm.PeriodLock>();
        map.put("11", Mcm2006uForm.periodLock(support, timestamp));
        return map;
    }
    @Test void supportIdChangedWithSameTimestampStopsBeforeAnyUpdate() {
        var repo = repo(snapshot("NEW0001", TIME));
        var service = service(repo);
        var stale = form(snapshot("OLD0001", TIME));
        var error = assertThrows(IllegalStateException.class, () -> service.save(stale, "test", "test"));
        assertEquals(Mcm2006uInputCheck.FWM_0009, error.getMessage());
        verify(repo).lock(CONTRACT);
        verify(repo).lockPeriodVersions(CONTRACT);
        verify(repo, never()).updateKeiyaku(any(), anyString());
    }
    @Test void otherPeriodTimestampChangeStopsBeforeUpdate() {
        var repo = repo(snapshot("OLD0001", Timestamp.valueOf("2026-10-29 12:00:01")));
        var service = service(repo);
        assertThrows(IllegalStateException.class, () -> service.save(form(snapshot("OLD0001", TIME)), "test", "test"));
        verify(repo, never()).updateKeiyaku(any(), anyString());
    }
    @Test void addedOrRemovedPeriodAndMissingSnapshotStopBeforeUpdate() {
        var repo = repo(Map.of());
        var service = service(repo);
        assertThrows(IllegalStateException.class, () -> service.save(form(snapshot("OLD0001", TIME)), "test", "test"));
        assertThrows(IllegalStateException.class, () -> service.save(form(null), "test", "test"));
        verify(repo, never()).updateKeiyaku(any(), anyString());
        var added = new LinkedHashMap<>(snapshot("OLD0001", TIME));
        added.put("12", Mcm2006uForm.periodLock("OTHER01", TIME));
        var addedRepo = repo(added);
        assertThrows(IllegalStateException.class, () -> service(addedRepo).save(form(snapshot("OLD0001", TIME)), "test", "test"));
        verify(addedRepo, never()).updateKeiyaku(any(), anyString());
    }
    @Test void identicalPeriodSnapshotAllowsNormalUpdatePath() {
        var matching = snapshot(null, null);
        var repo = repo(matching);
        when(repo.existsKeiyaku(CONTRACT)).thenReturn(true);
        var service = service(repo);
        var f = form(matching);
        assertDoesNotThrow(() -> service.save(f, "test", "test"));
        verify(repo).updateKeiyaku(f, "test");
    }
    @Test void newContractDoesNotRequireExistingPeriodSnapshot() {
        var repo = mock(Mcm2006uRepository.class);
        when(repo.nextKeiyakuId()).thenReturn(CONTRACT);
        var service = service(repo);
        var f = new Mcm2006uForm();
        // 入力チェックは別の試験範囲。ここでは新規登録経路への影響を確認する。
        f.setKeiyakuDt("2026/10/29");
        f.setShokaiKeiyakuDt("2026/10/29");
        assertDoesNotThrow(() -> service.save(f, "test", "test"));
        verify(repo).insertKeiyaku(f, "test");
        verify(repo, never()).lockPeriodVersions(any());
    }
    @Test void snapshotSurvivesSessionCopyAndNormalizesDatabaseIds() {
        var f = form(snapshot("OLD0001", TIME));
        var copy = CustomerScreenSupport.copy(f);
        assertEquals(f.getOriginalPeriodLocks(), copy.getOriginalPeriodLocks());
        assertEquals("11", Mcm2006uForm.periodLockKey(new BigDecimal("11.000")));
        assertNotEquals(Mcm2006uForm.periodLock(null, null), Mcm2006uForm.periodLock("", null));
    }
    @Test void loadedPeriodSnapshotIsIndependentOfLaterEdits() {
        var repo = repo(Map.of());
        var tab = new Mcm2006uForm.KikanTabForm();
        tab.setUkKikanId(new BigDecimal("11.00"));
        tab.setSupportId("OLD0001");
        tab.setOriginalPeriodLock(Mcm2006uForm.periodLock("OLD0001", TIME));
        when(repo.findKikanList(CONTRACT)).thenReturn(List.of(tab));
        var loaded = service(repo).load(CONTRACT);
        tab.setSupportId("NEW0001");
        assertEquals(snapshot("OLD0001", TIME), loaded.getOriginalPeriodLocks());
    }
    @Test void repositoryCapturesSupportAndTimestampFromSamePeriodRow() throws Exception {
        var jdbc = mock(JdbcTemplate.class);
        var repo = new Mcm2006uRepository();
        ReflectionTestUtils.setField(repo, "jdbc", jdbc);
        var row = mock(ResultSet.class);
        when(row.getBigDecimal("UK_KIKAN_ID")).thenReturn(BigDecimal.valueOf(11));
        when(row.getString("SUPPORT_ID")).thenReturn("OLD0001");
        when(row.getTimestamp("LASTUPDATE_DT")).thenReturn(TIME);
        when(jdbc.query(anyString(), any(Object[].class), any(RowMapper.class))).thenAnswer(inv -> {
            assertTrue(inv.getArgument(0, String.class).contains("LASTUPDATE_DT"));
            @SuppressWarnings("unchecked") RowMapper<Mcm2006uForm.KikanTabForm> mapper = inv.getArgument(2);
            return List.of(mapper.mapRow(row, 0));
        });
        var loaded = repo.findKikanList(CONTRACT);
        assertEquals(Mcm2006uForm.periodLock("OLD0001", TIME), loaded.get(0).getOriginalPeriodLock());
    }
    @Test void lockedPeriodQueryReturnsComparableDatabaseSnapshot() throws Exception {
        var jdbc = mock(JdbcTemplate.class);
        var repo = new Mcm2006uRepository();
        ReflectionTestUtils.setField(repo, "jdbc", jdbc);
        var row = mock(ResultSet.class);
        when(row.getBigDecimal("UK_KIKAN_ID")).thenReturn(new BigDecimal("11.000"));
        when(row.getString("SUPPORT_ID")).thenReturn("NEW0001");
        when(row.getTimestamp("LASTUPDATE_DT")).thenReturn(TIME);
        doAnswer(inv -> {
            RowCallbackHandler callback = inv.getArgument(1);
            callback.processRow(row);
            return null;
        }).when(jdbc).query(anyString(), any(RowCallbackHandler.class), any(Object[].class));
        assertEquals(snapshot("NEW0001", TIME), repo.lockPeriodVersions(CONTRACT));
    }
}
