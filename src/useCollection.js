import { useState, useEffect, useCallback, useRef } from "react";
import { supabase } from "./supabaseClient";

// Generic hook: keeps `rows` in sync with a Supabase table (initial fetch +
// realtime subscription), and exposes CRUD helpers. Every person running the
// app against the same Supabase project sees each other's changes live.
//
// Pass { withAudit: true } for tables that have created_by/updated_by columns
// — insert/bulkInsert/update will auto-fill them from the signed-in user,
// so callers never need to set them manually.
//
// `llave` es la columna que identifica cada fila (la llave primaria). Casi
// todas las tablas usan `id`; vehiculo_ubicaciones usa `codigo`. Se usa como
// desempate al paginar y para reconocer una fila en los eventos en vivo.
export function useCollection(table, orderBy = "created_at", options = {}) {
  const { withAudit = false, llave = "id" } = options;
  const [rows, setRows] = useState([]);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState(null);
  const rowsRef = useRef(rows);
  rowsRef.current = rows;

  useEffect(() => {
    let mounted = true;
    const k = (r) => r?.[llave];

    const fetchAll = async () => {
      // Supabase limita cada consulta a un máximo de filas por default (1000
      // en la mayoría de los proyectos) — con tablas grandes hace falta pedir
      // por páginas hasta traer todo, o las filas más recientes se quedan
      // fuera silenciosamente sin que se vea ningún error.
      //
      // El orden de las páginas tiene que ser TOTAL: ordenar solo por
      // created_at no basta, porque una importación masiva pone el mismo
      // created_at a cientos de filas (en transacciones hay 971 empatadas) y
      // entre filas empatadas Postgres no garantiza el mismo orden de una
      // consulta a otra. Si el corte de página cae dentro del empate, la
      // página 2 puede repetir filas de la 1 y saltarse otras: llegaban
      // transacciones duplicadas y faltaban otras. Los duplicados rompían la
      // tabla (dos filas con la misma key de React) y la búsqueda dejaba de
      // refrescarse. La llave primaria como desempate lo hace determinista.
      const PAGE_SIZE = 1000;
      let allRows = [];
      let from = 0;
      while (true) {
        let q = supabase.from(table).select("*").order(orderBy, { ascending: true });
        if (orderBy !== llave) q = q.order(llave, { ascending: true });
        const { data, error: err } = await q.range(from, from + PAGE_SIZE - 1);
        if (err) {
          if (!mounted) return;
          setError(err);
          setReady(true);
          return;
        }
        allRows = allRows.concat(data || []);
        if (!data || data.length < PAGE_SIZE) break;
        from += PAGE_SIZE;
      }
      // Red de seguridad: si una fila se movió de página mientras se leía
      // (alguien la editó en medio), que no entre dos veces.
      const vistos = new Set();
      allRows = allRows.filter((r) => (vistos.has(k(r)) ? false : (vistos.add(k(r)), true)));
      if (!mounted) return;
      setRows(allRows);
      setReady(true);
    };

    fetchAll();

    // La sesión de Supabase Auth puede terminar de establecerse DESPUÉS de que
    // este hook ya intentó su primera carga (que entonces sale sin token válido
    // y las políticas la bloquean). Cuando el estado de auth cambia — login,
    // sesión restaurada al abrir la pestaña, refresco de token — reintentamos.
    const { data: authListener } = supabase.auth.onAuthStateChange(() => {
      fetchAll();
    });

    const channel = supabase
      .channel(`realtime:${table}`)
      .on("postgres_changes", { event: "*", schema: "public", table }, (payload) => {
        setRows((current) => {
          if (payload.eventType === "INSERT") {
            if (current.some((r) => k(r) === k(payload.new))) return current;
            return [...current, payload.new];
          }
          if (payload.eventType === "UPDATE") {
            return current.map((r) => (k(r) === k(payload.new) ? payload.new : r));
          }
          if (payload.eventType === "DELETE") {
            return current.filter((r) => k(r) !== k(payload.old));
          }
          return current;
        });
      })
      .subscribe();

    return () => {
      mounted = false;
      authListener.subscription.unsubscribe();
      supabase.removeChannel(channel);
    };
  }, [table, orderBy, llave]);

  const currentUserId = useCallback(async () => {
    if (!withAudit) return null;
    const { data } = await supabase.auth.getUser();
    return data?.user?.id || null;
  }, [withAudit]);

  const insert = useCallback(async (record) => {
    let payload = record;
    if (withAudit) {
      const uid = await currentUserId();
      if (uid) payload = { ...record, created_by: record.created_by ?? uid, updated_by: uid };
    }
    const { data, error: err } = await supabase.from(table).insert(payload).select().single();
    if (err) { setError(err); throw err; }
    setRows((current) => (current.some((r) => r[llave] === data[llave]) ? current : [...current, data]));
    return data;
  }, [table, withAudit, currentUserId, llave]);

  const bulkInsert = useCallback(async (records) => {
    if (!records.length) return [];
    let payload = records;
    if (withAudit) {
      const uid = await currentUserId();
      if (uid) payload = records.map((r) => ({ ...r, created_by: r.created_by ?? uid, updated_by: uid }));
    }
    const { data, error: err } = await supabase.from(table).insert(payload).select();
    if (err) { setError(err); throw err; }
    // El evento en vivo de cada fila puede llegar antes que esta respuesta:
    // se agregan solo las que no estén ya.
    setRows((current) => {
      const ya = new Set(current.map((r) => r[llave]));
      return [...current, ...data.filter((r) => !ya.has(r[llave]))];
    });
    return data;
  }, [table, withAudit, currentUserId, llave]);

  const update = useCallback(async (id, patch) => {
    let payload = patch;
    if (withAudit) {
      const uid = await currentUserId();
      if (uid) payload = { ...patch, updated_by: uid };
    }
    const { data, error: err } = await supabase.from(table).update(payload).eq(llave, id).select().single();
    if (err) { setError(err); throw err; }
    setRows((current) => current.map((r) => (r[llave] === id ? data : r)));
    return data;
  }, [table, withAudit, currentUserId, llave]);

  const remove = useCallback(async (id) => {
    const { error: err } = await supabase.from(table).delete().eq(llave, id);
    if (err) { setError(err); throw err; }
    setRows((current) => current.filter((r) => r[llave] !== id));
  }, [table, llave]);

  const removeWhere = useCallback(async (column, values) => {
    if (!values.length) return;
    const { error: err } = await supabase.from(table).delete().in(column, values);
    if (err) { setError(err); throw err; }
    setRows((current) => current.filter((r) => !values.includes(r[column])));
  }, [table]);

  return { rows, ready, error, insert, bulkInsert, update, remove, removeWhere };
}
