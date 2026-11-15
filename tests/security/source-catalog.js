// Read-only public-schema metadata captured on 2026-11-06 from the user's Supabase project.
// Contains no table rows, credentials, personal data, or connection settings.
// Function owners/table grants/default ACLs are NOT fully restored; see acceptance report.
const functions = [
  {
    "name": "can_access_production_department",
    "arguments": "target_department_id uuid",
    "definition": "CREATE OR REPLACE FUNCTION public.can_access_production_department(target_department_id uuid)\n RETURNS boolean\n LANGUAGE sql\n STABLE SECURITY DEFINER\n SET search_path TO 'public'\nAS $function$\n  select exists (\n    select 1\n    from public.employees e\n    where e.auth_user_id = auth.uid()\n      and e.is_active = true\n      and (\n        e.role = 'admin'\n        or (\n          e.role = 'department_admin'\n          and target_department_id is not null\n          and exists (\n            select 1\n            from public.departments d\n            where d.id = target_department_id\n              and d.leader_employee_id = e.id\n              and d.active = true\n          )\n        )\n      )\n  )\n$function$\n",
    "security_definer": true,
    "settings": "{search_path=public}"
  },
  {
    "name": "current_app_employee_id",
    "arguments": "",
    "definition": "CREATE OR REPLACE FUNCTION public.current_app_employee_id()\n RETURNS uuid\n LANGUAGE sql\n STABLE SECURITY DEFINER\n SET search_path TO 'public'\nAS $function$\n  select e.id\n  from public.employees e\n  where e.auth_user_id = auth.uid()\n    and e.is_active = true\n  limit 1\n$function$\n",
    "security_definer": true,
    "settings": "{search_path=public}"
  },
  {
    "name": "current_app_employee_role",
    "arguments": "",
    "definition": "CREATE OR REPLACE FUNCTION public.current_app_employee_role()\n RETURNS text\n LANGUAGE sql\n STABLE SECURITY DEFINER\n SET search_path TO 'public'\nAS $function$\n  select e.role\n  from public.employees e\n  where e.auth_user_id = auth.uid()\n    and e.is_active = true\n  limit 1\n$function$\n",
    "security_definer": true,
    "settings": "{search_path=public}"
  },
  {
    "name": "current_employee_id",
    "arguments": "",
    "definition": "CREATE OR REPLACE FUNCTION public.current_employee_id()\n RETURNS uuid\n LANGUAGE sql\n STABLE SECURITY DEFINER\n SET search_path TO 'public'\nAS $function$\r\n  select e.id\r\n  from public.employees e\r\n  where e.auth_user_id = auth.uid()\r\n    and e.is_active = true\r\n  limit 1\r\n$function$\n",
    "security_definer": true,
    "settings": "{search_path=public}"
  },
  {
    "name": "is_admin",
    "arguments": "",
    "definition": "CREATE OR REPLACE FUNCTION public.is_admin()\n RETURNS boolean\n LANGUAGE sql\n STABLE SECURITY DEFINER\n SET search_path TO 'public'\nAS $function$\r\n  select exists (\r\n    select 1\r\n    from public.employees e\r\n    where e.auth_user_id = auth.uid()\r\n      and e.is_active = true\r\n      and e.role = 'admin'\r\n  )\r\n$function$\n",
    "security_definer": true,
    "settings": "{search_path=public}"
  },
  {
    "name": "set_updated_at",
    "arguments": "",
    "definition": "CREATE OR REPLACE FUNCTION public.set_updated_at()\n RETURNS trigger\n LANGUAGE plpgsql\nAS $function$\r\nbegin\r\n  new.updated_at = now();\r\n  return new;\r\nend;\r\n$function$\n",
    "security_definer": false,
    "settings": ""
  }
];
const standaloneIndexes = [
  {
    "table_name": "employees",
    "name": "employees_personnel_no_unique",
    "definition": "CREATE UNIQUE INDEX employees_personnel_no_unique ON public.employees USING btree (personnel_no) WHERE ((personnel_no IS NOT NULL) AND (personnel_no <> ''::text))"
  },
  {
    "table_name": "employees",
    "name": "employees_slot_code_unique",
    "definition": "CREATE UNIQUE INDEX employees_slot_code_unique ON public.employees USING btree (slot_code) WHERE (slot_code IS NOT NULL)"
  },
  {
    "table_name": "order_events",
    "name": "order_events_tool_idx",
    "definition": "CREATE INDEX order_events_tool_idx ON public.order_events USING btree (tool_id, created_at DESC)"
  },
  {
    "table_name": "planner_absence_replacements",
    "name": "planner_absence_replacements_shift_idx",
    "definition": "CREATE INDEX planner_absence_replacements_shift_idx ON public.planner_absence_replacements USING btree (shift_date, shift_id)"
  },
  {
    "table_name": "planner_absence_replacements",
    "name": "planner_absence_replacements_source_idx",
    "definition": "CREATE INDEX planner_absence_replacements_source_idx ON public.planner_absence_replacements USING btree (source_type, source_id)"
  },
  {
    "table_name": "planner_absences",
    "name": "planner_absences_date_idx",
    "definition": "CREATE INDEX planner_absences_date_idx ON public.planner_absences USING btree (absence_date)"
  },
  {
    "table_name": "planner_assignments",
    "name": "planner_assignments_shift_date_idx",
    "definition": "CREATE INDEX planner_assignments_shift_date_idx ON public.planner_assignments USING btree (shift_date)"
  },
  {
    "table_name": "planner_availability",
    "name": "planner_availability_shift_idx",
    "definition": "CREATE INDEX planner_availability_shift_idx ON public.planner_availability USING btree (shift_date, shift_id)"
  },
  {
    "table_name": "planner_saturday_requests",
    "name": "planner_saturday_requests_shift_idx",
    "definition": "CREATE INDEX planner_saturday_requests_shift_idx ON public.planner_saturday_requests USING btree (shift_date, shift_id)"
  },
  {
    "table_name": "planner_shift_cancellations",
    "name": "planner_shift_cancellations_date_idx",
    "definition": "CREATE INDEX planner_shift_cancellations_date_idx ON public.planner_shift_cancellations USING btree (shift_date)"
  },
  {
    "table_name": "planner_sick_leaves",
    "name": "planner_sick_leaves_user_idx",
    "definition": "CREATE INDEX planner_sick_leaves_user_idx ON public.planner_sick_leaves USING btree (user_name, from_date, to_date)"
  },
  {
    "table_name": "planner_swaps",
    "name": "planner_swaps_dates_idx",
    "definition": "CREATE INDEX planner_swaps_dates_idx ON public.planner_swaps USING btree (start_date, end_date)"
  },
  {
    "table_name": "planner_vacations",
    "name": "planner_vacations_user_idx",
    "definition": "CREATE INDEX planner_vacations_user_idx ON public.planner_vacations USING btree (user_name, from_date, to_date)"
  },
  {
    "table_name": "production_order_checklist",
    "name": "production_order_checklist_order_idx",
    "definition": "CREATE INDEX production_order_checklist_order_idx ON public.production_order_checklist USING btree (order_id)"
  },
  {
    "table_name": "production_order_employees",
    "name": "production_order_employees_order_idx",
    "definition": "CREATE INDEX production_order_employees_order_idx ON public.production_order_employees USING btree (order_id)"
  },
  {
    "table_name": "production_order_history",
    "name": "production_order_history_order_created_idx",
    "definition": "CREATE INDEX production_order_history_order_created_idx ON public.production_order_history USING btree (order_id, created_at DESC)"
  },
  {
    "table_name": "production_order_stations",
    "name": "production_order_stations_order_idx",
    "definition": "CREATE INDEX production_order_stations_order_idx ON public.production_order_stations USING btree (order_id)"
  },
  {
    "table_name": "production_orders",
    "name": "production_orders_department_status_idx",
    "definition": "CREATE INDEX production_orders_department_status_idx ON public.production_orders USING btree (department_id, status)"
  },
  {
    "table_name": "production_orders",
    "name": "production_orders_machine_idx",
    "definition": "CREATE INDEX production_orders_machine_idx ON public.production_orders USING btree (machine_id)"
  },
  {
    "table_name": "production_station_counts",
    "name": "production_station_counts_order_station_idx",
    "definition": "CREATE INDEX production_station_counts_order_station_idx ON public.production_station_counts USING btree (order_id, station_id)"
  },
  {
    "table_name": "production_station_events",
    "name": "production_station_events_order_created_idx",
    "definition": "CREATE INDEX production_station_events_order_created_idx ON public.production_station_events USING btree (order_id, created_at DESC)"
  },
  {
    "table_name": "tool_journal",
    "name": "tool_journal_tool_idx",
    "definition": "CREATE INDEX tool_journal_tool_idx ON public.tool_journal USING btree (tool_id, created_at DESC)"
  },
  {
    "table_name": "tools",
    "name": "tools_label_idx",
    "definition": "CREATE INDEX tools_label_idx ON public.tools USING btree (label)"
  },
  {
    "table_name": "tools",
    "name": "tools_manufacturer_idx",
    "definition": "CREATE INDEX tools_manufacturer_idx ON public.tools USING btree (manufacturer)"
  },
  {
    "table_name": "tools",
    "name": "tools_stock_idx",
    "definition": "CREATE INDEX tools_stock_idx ON public.tools USING btree (stock, min_stock)"
  }
];
module.exports={functions,standaloneIndexes};
