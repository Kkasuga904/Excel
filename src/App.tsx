import React, { useEffect, useMemo, useState } from 'react';
import TaskPane from './taskpane/taskpane';

type HealthState = {
  state: 'loading' | 'success' | 'error';
  message: string;
};

const defaultHealthMessage = 'サーバーへの接続を確認しています...';

const App: React.FC = () => {
  const apiBase = useMemo(() => process.env.REACT_APP_API_BASE_URL ?? '', []);
  const [healthStatus, setHealthStatus] = useState<HealthState>({
    state: 'loading',
    message: defaultHealthMessage
  });

  useEffect(() => {
    const abortController = new AbortController();
    const checkHealth = async () => {
      setHealthStatus({ state: 'loading', message: defaultHealthMessage });

      try {
        const normalizedBase = apiBase.replace(/\/+$/, '');
        const endpoint = normalizedBase ? `${normalizedBase}/health` : '/health';
        const response = await fetch(endpoint, {
          signal: abortController.signal
        });

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const payload: { timestamp?: string } | undefined = await response
          .json()
          .catch(() => undefined);

        const timestamp = payload?.timestamp
          ? new Date(payload.timestamp).toLocaleString('ja-JP')
          : undefined;

        setHealthStatus({
          state: 'success',
          message: timestamp
            ? `✓ サーバーに接続しました（${timestamp}）`
            : '✓ サーバーに接続しました'
        });
      } catch (error) {
        const detail = error instanceof Error ? error.message : 'unknown error';
        setHealthStatus({
          state: 'error',
          message: `✗ サーバーに接続できません。環境変数 REACT_APP_API_BASE_URL を確認してください。（${detail}）`
        });
      }
    };

    void checkHealth();

    return () => {
      abortController.abort();
    };
  }, [apiBase]);

  return <TaskPane healthStatus={healthStatus} />;
};

export default App;
