function appConfig() {
  return {
    DB_UTAMA: "1VOMGDOXD7C1iONXgdpOzxyc4Xmzueow-yQHsAwPQfUw",
    RAMADAN_PERIOD: { START: "2026-02-22", END: "2026-03-22" },
    NORMAL_SHIFT: {
      MON_THU: { START: "07:30:00", END: "16:30:00", REST: "12:00-13:00" },
      FRI:     { START: "07:30:00", END: "17:00:00", REST: "12:00-13:00" }
    },
    RAMADAN_SHIFT: {
      MON_THU: { START: "07:30:00", END: "16:00:00", REST: "12:00-12:30" },
      FRI:     { START: "07:30:00", END: "16:30:00", REST: "12:00-13:00" }
    }
  };
}
