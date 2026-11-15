import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  CircularProgress,
  Divider,
  List,
  ListItem,
  ListItemText,
  Paper,
  Stack,
  TextField,
  Typography,
} from "@mui/material";
import { useTranslation } from "react-i18next";
import {
  fetchDefenseQueue,
  registerAndCheck,
  registerForDefenseQueue,
  shiftDefenseQueue,
  withdrawFromDefenseQueue,
} from "../../api";
import { alpha } from "@mui/material/styles";

const STATUS_CALLED = "вызван";

const STATUS_BG = {
  "сдал": (theme) => alpha(theme.palette.success.main, 0.1),
  "не_сдал": (theme) => alpha(theme.palette.error.main, 0.08),
  "нет": (theme) => alpha(theme.palette.text.primary, 0.06),
  "отозван": (theme) => alpha(theme.palette.text.primary, 0.06),
  "не_дошло": (theme) => alpha(theme.palette.text.primary, 0.06),
  [STATUS_CALLED]: (theme) => alpha(theme.palette.info.main, 0.14),
};

const normalizeQueue = (data) => {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.queue)) return data.queue;
  if (Array.isArray(data?.entries)) return data.entries;
  if (Array.isArray(data?.rows)) return data.rows;
  return [];
};

const validateStatus = (data) => {
  if (!data || typeof data !== "object" || typeof data.in_queue !== "boolean") {
    return null;
  }
  return data;
};

const formatSessionMoment = (value, language) => {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat(language, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
};

const sessionLabel = (session, language, t) => {
  if (!session) return t("queue.noNextSession");
  const start = formatSessionMoment(session.starts_at, language);
  const end = formatSessionMoment(session.ends_at, language);
  const timeRange = start && end ? `${start} – ${end}` : start || end;
  return [timeRange, session.room].filter(Boolean).join(" · ") || t("queue.noNextSession");
};

const entryName = (entry) =>
  entry.team_title ||
  entry.team?.title ||
  entry.fio ||
  entry.student_name ||
  entry.name ||
  entry.title ||
  "";

const entryMeta = (entry, t) => {
  const github = entry.github ? String(entry.github).trim() : "";
  const signupTime = entry.signup_time ? String(entry.signup_time).trim() : "";
  if (!github && !signupTime) return null;
  const parts = [];
  if (github) parts.push(`${t("queue.githubLabel")}: ${github}`);
  if (signupTime) parts.push(`${t("queue.signupTimeLabel")}: ${signupTime}`);
  else parts.push(`${t("queue.signupTimeLabel")}: ${t("queue.signupTimeEmpty")}`);
  return parts.join(" · ");
};

// eslint-disable-next-line react/prop-types
export function QueueSection({ courseId, groupId, labId, formData }) {
  const { t, i18n } = useTranslation();
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const [notice, setNotice] = useState(null);
  const [shiftBy, setShiftBy] = useState("1");

  // Guard against out-of-order responses: every load gets a sequence
  // number and a response that arrives after a newer request has started
  // (or after unmount) is discarded instead of overwriting fresh state.
  const requestSeq = useRef(0);
  const isMountedRef = useRef(false);

  // П.14: один GET /queue возвращает и список, и персональное состояние
  // (me) — отдельный GET /queue/status больше не нужен.
  const loadQueue = useCallback(
    async ({ showLoading = false } = {}) => {
      const seq = ++requestSeq.current;
      if (showLoading) {
        setLoading(true);
      }
      setError(null);
      try {
        const resp = await fetchDefenseQueue(courseId, groupId, labId);
        if (seq !== requestSeq.current) return; // stale response - discard
        setData(resp);
      } catch (requestError) {
        if (seq !== requestSeq.current) return; // stale failure - discard
        setData(null);
        setError(requestError);
      } finally {
        if (seq === requestSeq.current && isMountedRef.current) {
          setLoading(false);
        }
      }
    },
    [courseId, groupId, labId]
  );

  const reload = useCallback(() => loadQueue({ showLoading: false }), [loadQueue]);

  useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  useEffect(() => {
    loadQueue({ showLoading: true });
  }, [loadQueue]);

  // П.14: поллинг 15 c, пауза на скрытой вкладке, обновление при возврате.
  useEffect(() => {
    const intervalId = setInterval(() => {
      if (!document.hidden) {
        loadQueue();
      }
    }, 15000);
    const onVisibilityChange = () => {
      if (!document.hidden) {
        loadQueue();
      }
    };
    document.addEventListener("visibilitychange", onVisibilityChange);
    return () => {
      clearInterval(intervalId);
      document.removeEventListener("visibilitychange", onVisibilityChange);
    };
  }, [loadQueue]);

  // Производные состояния из одного ответа /queue (П.14).
  const me = data?.me ?? null;
  const sessionRequired = me === null;
  const inQueue = me?.in_queue === true;
  const called = me?.called === true;
  const position = me?.position ?? null;
  const alreadyPassed = false; // проверяется при регистрации (register -> ALREADY_PASSED)
  const session = data?.session ?? null;
  const nextSession = data?.next_session ?? null;
  const queueOpen = data?.queue_open === true;
  const queueRows = useMemo(() => data?.queue ?? [], [data]);
  const teamLab = data?.is_team_lab === true;
  const team = data?.team ?? null;

  const handleJoinQueue = async () => {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      if (sessionRequired && formData) {
        await registerAndCheck(courseId, groupId, formData);
      }
      await registerForDefenseQueue(courseId, groupId, labId);
      setNotice(t("queue.registered"));
      await reload();
    } catch (requestError) {
      setError(requestError);
    } finally {
      setBusy(false);
    }
  };

  const handleWithdraw = async () => {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      await withdrawFromDefenseQueue(courseId, groupId, labId);
      setNotice(t("queue.withdrawn"));
      await reload();
    } catch (requestError) {
      setError(requestError);
    } finally {
      setBusy(false);
    }
  };

  const handleShift = async () => {
    setBusy(true);
    setError(null);
    setNotice(null);
    try {
      await shiftDefenseQueue(courseId, groupId, labId, Number(shiftBy), sessionNo);
      setNotice(t("queue.shifted"));
      await reload();
    } catch (requestError) {
      setError(requestError);
    } finally {
      setBusy(false);
    }
  };

  const nextSessionLabel = useMemo(
    () => sessionLabel(nextSession, i18n.language, t),
    [i18n.language, nextSession, t]
  );
  const currentSessionLabel = useMemo(
    () => sessionLabel(session, i18n.language, t),
    [i18n.language, session, t]
  );

  const errorMessage = error
    ? error.code
      ? t(`queue.errors.${error.code}`, {
          defaultValue: error.message || t("queue.errors.generic"),
        })
      : error.message || t("queue.errors.generic")
    : null;

  const teamMembers = Array.isArray(team?.members) ? team.members : [];
  const validShift = /^[1-9]\d*$/.test(shiftBy) && Number.isSafeInteger(Number(shiftBy));
  const formReady = formData && formData.surname && formData.name && formData.github;

  return (
    <Paper component="section" variant="outlined" sx={{ mt: 3, p: { xs: 2, sm: 3 } }}>
      <Stack spacing={2}>
        <Box>
          <Typography variant="h6">{t("queue.title")}</Typography>
          <Typography variant="body2" color="text.secondary">
            {t("queue.description")}
          </Typography>
        </Box>

        {loading && (
          <Stack direction="row" spacing={1} alignItems="center" role="status">
            <CircularProgress size={18} />
            <Typography variant="body2">{t("queue.loading")}</Typography>
          </Stack>
        )}

        {!loading && errorMessage && <Alert severity="error">{errorMessage}</Alert>}
        {notice && <Alert severity="success">{notice}</Alert>}

        {!loading && (
          <>
            {teamLab && team && (
              <Paper variant="outlined" sx={{ p: 2 }}>
                <Typography variant="subtitle2">{t("queue.teamTitle")}</Typography>
                <Typography>{team.title || team.name || team.slug}</Typography>
                {teamMembers.length > 0 && (
                  <Typography variant="body2" color="text.secondary">
                    {t("queue.teamMembers")}:{" "}
                    {teamMembers
                      .map((member) =>
                        typeof member === "string"
                          ? member
                          : member.name || member.fio || member.username || member.github
                      )
                      .filter(Boolean)
                      .join(", ")}
                  </Typography>
                )}
              </Paper>
            )}

            {inQueue && called && (
              <Alert severity="info">{t("queue.calledBanner")}</Alert>
            )}

            {inQueue && !called && (
              <Alert severity="success">
                {t("queue.position", {
                  position: position,
                  total: me?.queue_size ?? position,
                })}
              </Alert>
            )}

            {!inQueue && !called && !queueOpen && (
              <Alert severity="warning">
                {t("queue.closed", { nextSession: nextSessionLabel })}
              </Alert>
            )}

            {!inQueue && !called && queueOpen && (
              <Button
                variant="contained"
                color="primary"
                size="large"
                disabled={busy || !formReady}
                onClick={handleJoinQueue}
              >
                {t(busy ? "queue.working" : "queue.register")}
              </Button>
            )}

            {inQueue && !called && (
              <Stack direction={{ xs: "column", sm: "row" }} spacing={1} alignItems="flex-start">
                <Button
                  variant="outlined"
                  color="error"
                  disabled={busy}
                  onClick={handleWithdraw}
                >
                  {t(busy ? "queue.working" : "queue.withdraw")}
                </Button>
                <Stack direction="row" spacing={1} alignItems="center">
                  <TextField
                    type="number"
                    size="small"
                    label={t("queue.shiftBy")}
                    value={shiftBy}
                    onChange={(event) => setShiftBy(event.target.value)}
                    inputProps={{ min: 1, step: 1 }}
                    error={shiftBy !== "" && !validShift}
                    sx={{ width: 150 }}
                  />
                  <Button
                    variant="outlined"
                    disabled={busy || !validShift}
                    onClick={handleShift}
                  >
                    {t(busy ? "queue.working" : "queue.shift")}
                  </Button>
                </Stack>
              </Stack>
            )}

            {session && (
              <Typography variant="body2" color="text.secondary">
                {t("queue.session")}: {currentSessionLabel}
              </Typography>
            )}

            <Divider />
            <Box>
              <Stack
                direction="row"
                alignItems="center"
                justifyContent="space-between"
                sx={{ mb: 1 }}
              >
                <Typography variant="subtitle1">{t("queue.listTitle")}</Typography>
                <Button size="small" disabled={busy || loading} onClick={reload}>
                  {t("queue.refresh")}
                </Button>
              </Stack>
              {queueRows.length === 0 ? (
                <Typography variant="body2" color="text.secondary">
                  {t("queue.empty")}
                </Typography>
              ) : (
                <List disablePadding>
                  {queueRows.map((entry, index) => {
                    const isCalled = entry.status === STATUS_CALLED;
                    const position = entry.position ?? (isCalled ? "—" : index + 1);
                    return (
                      <ListItem
                        key={`${entry.fio}-${entry.lab}-${index}`}
                        divider={index < queueRows.length - 1}
                        disableGutters
                        sx={
                          STATUS_BG[entry.status]
                            ? { bgcolor: STATUS_BG[entry.status] }
                            : undefined
                        }
                        secondaryAction={
                          entry.status && (
                            <Chip
                              size="small"
                              color={isCalled ? "info" : "default"}
                              label={t(`queue.statuses.${entry.status}`, {
                                defaultValue: entry.status,
                              })}
                            />
                          )
                        }
                      >
                        <ListItemText
                          primary={
                            <Stack direction="row" spacing={1} alignItems="center">
                              <Chip size="small" label={position} />
                              <Typography component="span" variant="body2">
                                {entryName(entry) || t("queue.unnamedEntry")}
                              </Typography>
                              {/* One queue per pair: show which lab each
                                  participant is defending (issue #62 review). */}
                              {entry.lab ? (
                                <Chip size="small" variant="outlined" label={entry.lab} />
                              ) : null}
                            </Stack>
                          }
                          secondary={entryMeta(entry, t) || undefined}
                          secondaryTypographyProps={{ variant: "caption", color: "text.secondary" }}
                        />
                      </ListItem>
                    );
                  })}
                </List>
              )}
            </Box>
          </>
        )}
      </Stack>
    </Paper>
  );
}
