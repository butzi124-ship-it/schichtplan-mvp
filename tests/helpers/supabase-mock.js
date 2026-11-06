const rowsByTable = {
  employees: [
    {
      id: "employee-admin",
      auth_user_id: "auth-admin",
      display_name: "Admin",
      role: "admin",
      employee_type: "admin",
      slot_code: null,
      is_active: true,
      first_name: "Mido",
      last_name: "Admin",
      personnel_no: "A-1",
      department: "",
      color_key: "slate",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "employee-one",
      auth_user_id: null,
      display_name: "Lavdrim",
      role: "employee",
      employee_type: "core",
      slot_code: "A",
      is_active: true,
      first_name: "Lavdrim",
      last_name: "",
      personnel_no: "100",
      department: "",
      color_key: "blue",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  shift_definitions: [
    {
      id: "shift-early",
      name: "Frühschicht",
      shift_key: "early",
      start_time: "06:00",
      end_time: "14:00",
      requires_time: true,
      active: true,
    },
    {
      id: "shift-day",
      name: "Tagschicht",
      shift_key: "day",
      start_time: null,
      end_time: null,
      requires_time: false,
      active: true,
    },
  ],
  departments: [
    {
      id: "department-one",
      name: "Fertigung",
      code: "FERT",
      leader_employee_id: "employee-admin",
      active: true,
    },
  ],
  production_machines: [
    {
      id: "machine-one",
      name: "Maschine 50",
      machine_code: "M50",
      department_id: "department-one",
      active: true,
    },
    {
      id: "machine-two",
      name: "Maschine 51",
      machine_code: "M51",
      department_id: "department-one",
      active: true,
    },
    {
      id: "machine-over",
      name: "Maschine 52",
      machine_code: "M52",
      department_id: "department-one",
      active: true,
    },
    {
      id: "machine-chain",
      name: "Maschine 53",
      machine_code: "M53",
      department_id: "department-one",
      active: true,
    },
  ],
  production_orders: [
    {
      id: "order-one",
      machine_id: "machine-one",
      department_id: "department-one",
      ba_number: "BA-100",
      article_number: "ART-100",
      ba_quantity: 100,
      target_quantity: 2,
      pallet_count: 2,
      pieces_per_pallet: 50,
      use_chain_logic: false,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "order-two",
      machine_id: "machine-two",
      department_id: "department-one",
      ba_number: "BA-200",
      article_number: "ART-200",
      ba_quantity: 50,
      target_quantity: 50,
      pallet_count: 1,
      pieces_per_pallet: 50,
      use_chain_logic: false,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "order-over",
      machine_id: "machine-over",
      department_id: "department-one",
      ba_number: "BA-300",
      article_number: "ART-300",
      ba_quantity: 10,
      target_quantity: 2,
      pallet_count: 1,
      pieces_per_pallet: 10,
      use_chain_logic: false,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "order-chain",
      machine_id: "machine-chain",
      department_id: "department-one",
      ba_number: "BA-400",
      article_number: "ART-400",
      ba_quantity: 10,
      target_quantity: 2,
      pallet_count: 1,
      pieces_per_pallet: 10,
      use_chain_logic: true,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_order_stations: [
    {
      id: "station-one",
      order_id: "order-one",
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: "10",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "station-over",
      order_id: "order-over",
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: "10",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 1,
      clarify_total: 0,
      scrap_lifetime: 1,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "station-chain-one",
      order_id: "order-chain",
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: "10",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "station-chain-two",
      order_id: "order-chain",
      station_no: 2,
      name: "Spannung 2",
      lock_name: false,
      op_number: "20",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_order_employees: [
    {
      id: "entry-over",
      order_id: "order-over",
      employee_id: "employee-one",
      role: "worker",
      active: true,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "entry-chain",
      order_id: "order-chain",
      employee_id: "employee-one",
      role: "worker",
      active: true,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_checklist_templates: [
    "Stückzahl gezählt?",
    "Auftrag fertig gemeldet?",
    "Lagerort eingetragen?",
    "Hilfsmittel eingeräumt?",
    "Werkzeuge eingeräumt?",
    "Teile eingeölt?",
    "Doku ergänzt?",
    "Excel eingetragen?",
    "Doku-Fehler notiert?",
    "Doku abgelegt?",
  ].map((label, index) => ({
    id: `template-${index + 1}`,
    item_key: `item_${index + 1}`,
    item_label: label,
    active: true,
    sort_order: index + 1,
    created_at: "2026-01-01T00:00:00Z",
    updated_at: "2026-01-01T00:00:00Z",
  })),
  production_order_checklist: [],
  production_order_history: [
    {
      id: "history-completed",
      order_id: "order-one",
      station_id: null,
      employee_id: "employee-admin",
      history_type: "order_completed",
      qty: 0,
      payload: {
        ba_number: "BA-100",
        good_total: 1,
        scrap_total: 1,
        clarify_total: 0,
        remaining: 0,
      },
      note: "Historie Test",
      created_at: "2026-01-01T00:04:00Z",
    },
  ],
  production_station_counts: [
    {
      id: "count-over",
      order_id: "order-over",
      station_id: "station-over",
      order_employee_id: "entry-over",
      good_qty: 2,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "count-chain-one",
      order_id: "order-chain",
      station_id: "station-chain-one",
      order_employee_id: "entry-chain",
      good_qty: 2,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_station_events: [
    {
      id: "event-good",
      order_id: "order-one",
      station_id: "station-one",
      order_employee_id: "",
      employee_id: "employee-one",
      event_type: "good",
      qty: 1,
      qa_cause_id: null,
      note: "Gutteil Test",
      created_at: "2026-01-01T00:02:00Z",
    },
    {
      id: "event-scrap",
      order_id: "order-one",
      station_id: "station-one",
      order_employee_id: "",
      employee_id: "employee-admin",
      event_type: "scrap",
      qty: 1,
      qa_cause_id: "qa-cause-human",
      note: "Ausschuss Test",
      created_at: "2026-01-01T00:03:00Z",
    },
  ],
  production_qa_causes: [
    {
      id: "qa-cause-human",
      group_id: "human",
      group_label: "Mensch",
      reason_id: "training",
      reason_label: "Einweisung",
      active: true,
      sort_order: 1,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "qa-cause-machine",
      group_id: "machine",
      group_label: "Maschine",
      reason_id: "setup",
      reason_label: "Einrichtung",
      active: true,
      sort_order: 2,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_counts: [],
  tool_materials: [],
  tools: [
    {
      id: "tool-one",
      t_number: "133",
      label: "Schaftfräser",
      diameter: "10",
      corner_radius: "",
      article_no: "ART-133",
      holder: "HSK 63",
      shelf: "12N",
      stock: 4,
      min_stock: 1,
      optimal_stock: 4,
      manufacturer: "Test",
      ordered: false,
      ordered_qty: 0,
      insert_tool: false,
      insert_edges: 0,
      insert_radius: null,
      is_borrowed: false,
      borrowed_to: null,
      borrowed_at: null,
      material_id: null,
    },
  ],
  tool_journal: [],
  planner_tasks: [],
  planner_assignments: [],
  planner_absences: [],
  planner_vacations: [],
  planner_sick_leaves: [],
  planner_availability: [],
  planner_swaps: [],
  planner_shift_cancellations: [],
  planner_saturday_requests: [],
  planner_absence_replacements: [],
};

function installSupabaseMock(rowsByTableArg) {
  const rowsByTable = rowsByTableArg;
  window.__SUPABASE_MOCK_ROWS = rowsByTable;
  window.__SUPABASE_WRITE_LOG = [];
  window.__SUPABASE_MOCK_ERRORS = {};
  window.TEST_ADMIN_EMAIL = "admin@example.test";
  window.TEST_EMPLOYEE_EMAIL = "employee@example.test";

  function consumeMockError(table, action) {
    const tableErrors = window.__SUPABASE_MOCK_ERRORS?.[table];
    const error = tableErrors?.[action];
    if (error) {
      delete tableErrors[action];
      return error;
    }
    return null;
  }

  class QueryMock {
    constructor(table) {
      this.table = table;
      this.filters = [];
      this.useSingle = false;
      this.useMaybeSingle = false;
      this.pendingUpdate = null;
      this.pendingDelete = false;
    }

    select() { return this; }
    order() { return this; }
    limit() { return this; }
    insert(payload) {
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "insert", payload });
      const mockError = consumeMockError(this.table, "insert");
      if (mockError) return Promise.resolve({ data: null, error: mockError });
      const rows = Array.isArray(payload) ? payload : [payload];
      const tableRows = rowsByTable[this.table] || [];
      rowsByTable[this.table] = tableRows;
      const created = rows.map((row, index) => ({
        id: row.id || `${this.table}-${Date.now()}-${index}`,
        created_at: row.created_at || new Date().toISOString(),
        updated_at: row.updated_at || new Date().toISOString(),
        ...row,
      }));
      tableRows.push(...created);
      return Promise.resolve({ data: created, error: null });
    }
    update(payload) {
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "update", payload });
      this.pendingUpdate = payload || {};
      return this;
    }
    delete() {
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "delete" });
      this.pendingDelete = true;
      return this;
    }
    upsert(payload) { return this.insert(payload); }
    single() { this.useSingle = true; return this; }
    maybeSingle() { this.useMaybeSingle = true; return this; }
    eq(column, value) { this.filters.push({ column, value }); return this; }

    result() {
      const tableRows = rowsByTable[this.table] || [];
      let data = [...tableRows];
      for (const { column, value } of this.filters) {
        data = data.filter((row) => String(row[column]) === String(value));
      }
      const mockAction = this.pendingUpdate ? "update" : this.pendingDelete ? "delete" : "select";
      const mockError = consumeMockError(this.table, mockAction);
      if (mockError) return { data: null, error: mockError };
      if (this.pendingUpdate) {
        data.forEach((row) => Object.assign(row, this.pendingUpdate));
        return { data, error: null };
      }
      if (this.pendingDelete) {
        rowsByTable[this.table] = tableRows.filter((row) => !data.includes(row));
        return { data, error: null };
      }
      if (this.useSingle || this.useMaybeSingle) {
        return { data: data[0] || null, error: null };
      }
      return { data, error: null };
    }

    then(resolve, reject) {
      const hold = window.__SUPABASE_MOCK_HOLD;
      if (hold?.table === this.table && this.pendingUpdate) {
        return new Promise((release) => {
          hold.release = () => release(this.result());
        }).then(resolve, reject);
      }
      return Promise.resolve(this.result()).then(resolve, reject);
    }
  }

  window.supabase = {
    createClient: () => ({
      from: (table) => new QueryMock(table),
      auth: {
        getSession: async () => ({ data: { session: null }, error: null }),
        getUser: async () => ({ data: { user: null }, error: null }),
        signInWithPassword: async () => ({ data: { user: null }, error: null }),
        signOut: async () => ({ error: null }),
      },
      channel: () => ({
        on() { return this; },
        subscribe(callback) {
          if (typeof callback === "function") callback("SUBSCRIBED");
          return this;
        },
      }),
      removeChannel: async () => null,
    }),
  };
}

module.exports = { rowsByTable, installSupabaseMock };
