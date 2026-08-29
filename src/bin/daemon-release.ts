import { Command } from 'commander';
import { releaseDaemonAt, type StopResult } from '../mcp/daemon-registry';

/** Testable command handler behind `daemon stop --path`. */
export async function runDaemonStop(
  root: string,
  release: (path: string) => Promise<StopResult> = releaseDaemonAt,
): Promise<string> {
  return JSON.stringify(await release(root));
}

/** Register the non-interactive parser path without importing the full CLI entrypoint. */
export function registerDaemonStopCommand(
  parent: Command,
  run: (path: string) => Promise<string> = runDaemonStop,
  output: (value: string) => void = console.log,
): Command {
  return parent
    .command('stop')
    .description('Release the CodeGraph daemon and index locks for one explicit project root')
    .requiredOption('-p, --path <path>', 'Canonical project root to release')
    .action(async (options: { path: string }) => output(await run(options.path)));
}
