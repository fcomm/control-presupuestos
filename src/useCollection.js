import { useState, useEffect, useCallback, useRef } from "react";
import { supabase } from "./supabaseClient";

// Generic hook: keeps `rows` in sync with a Supabase table (initial fetch +
// realtime subscription), and exposes CRUD helpers. Every person running the
// app against the same Supabase project sees each other's changes live.
export function useCollection(table, orderBy = "created_at") {
  const [rows, setRows] = useState([]);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState(null);
  const rowsRef = useRef(rows);
  rowsRef.current = rows;

  useEffect(() => {
    let mounted = true;

    (async () => {
      const { data, error: err } = await supabase.from(table).select("*").order(orderBy, { ascending: true });
      if (!mounted) return;
      if (err) setError(err);
      else setRows(data || []);
      setReady(true);
    })();

    const channel = supabase
      .channel(`realtime:${table}`)
      .on("postgres_changes", { event: "*", schema: "public", table }, (payload) => {
        setRows((current) => {
          if (payload.eventType === "INSERT") {
            if (current.some((r) => r.id === payload.new.id)) return current;
            return [...current, payload.new];
          }
          if (payload.eventType === "UPDATE") {
            return current.map((r) => (r.id === payload.new.id ? payload.new : r));
          }
          if (payload.eventType === "DELETE") {
            return current.filter((r) => r.id !== payload.old.id);
          }
          return current;
        });
      })
      .subscribe();

    return () => {
      mounted = false;
      supabase.removeChannel(channel);
    };
  }, [table, orderBy]);

  const insert = useCallback(async (record) => {
    const { data, error: err } = await supabase.from(table).insert(record).select().single();
    if (err) { setError(err); throw err; }
    setRows((current) => (current.some((r) => r.id === data.id) ? current : [...current, data]));
    return data;
  }, [table]);

  const bulkInsert = useCallback(async (records) => {
    if (!records.length) return [];
    const { data, error: err } = await supabase.from(table).insert(records).select();
    if (err) { setError(err); throw err; }
    setRows((current) => [...current, ...data]);
    return data;
  }, [table]);

  const update = useCallback(async (id, patch) => {
    const { data, error: err } = await supabase.from(table).update(patch).eq("id", id).select().single();
    if (err) { setError(err); throw err; }
    setRows((current) => current.map((r) => (r.id === id ? data : r)));
    return data;
  }, [table]);

  const remove = useCallback(async (id) => {
    const { error: err } = await supabase.from(table).delete().eq("id", id);
    if (err) { setError(err); throw err; }
    setRows((current) => current.filter((r) => r.id !== id));
  }, [table]);

  const removeWhere = useCallback(async (column, values) => {
    if (!values.length) return;
    const { error: err } = await supabase.from(table).delete().in(column, values);
    if (err) { setError(err); throw err; }
    setRows((current) => current.filter((r) => !values.includes(r[column])));
  }, [table]);

  return { rows, ready, error, insert, bulkInsert, update, remove, removeWhere };
}
