// codegraph-launcher — Compatibility shim that resolves the canonical `codegraph`
// CLI to the user's installed @aroman22/codegraph-vba fork (npm or pnpm) without
// relying on broken symlinks to the fork's working tree.
//
// Search order (first match wins; CODEGRAPH_NPM_SHIM env var overrides all):
//   1. %CODEGRAPH_NPM_SHIM% (absolute path)
//   2. $USERPROFILE/node_modules/@aroman22/codegraph-vba/npm-shim.js   (user npm prefix)
//   3. $USERPROFILE/AppData/Roaming/npm/node_modules/@aroman22/codegraph-vba/npm-shim.js
//   4. $USERPROFILE/AppData/Local/pnpm/global/5/node_modules/@aroman22/codegraph-vba/npm-shim.js
//
// All searches use filepath.EvalSymlinks to skip dead symlinks, and every candidate
// is verified with os.Stat before use. node.exe is resolved via exec.LookPath so
// the user's PATH wins; no shell, no cmd.exe double-expansion. Exit codes are
// propagated verbatim so MCP server failures surface as MCP errors, not silent 0.
//
// Replacement target: replaces the legacy `codegraph.cmd` (which suffered the
// cmd.exe CALL %* second-expansion / metacharacter bug) and the broken batch
// adapter while keeping the original shim-installer architecture intact.
package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

func main() {
	npmShim, err := findNpmShim()
	if err != nil {
		fmt.Fprintf(os.Stderr, "codegraph compatibility error: %v\n", err)
		os.Exit(9009)
	}

	node, err := exec.LookPath("node")
	if err != nil {
		// On Windows, LookPath("node") still finds node.exe via PATHEXT, so this
		// error path is the same on every platform: the user has no Node runtime
		// on PATH at all.
		fmt.Fprintln(os.Stderr, "codegraph compatibility error: node was not found on PATH.")
		os.Exit(9009)
	}

	child := exec.Command(node, append([]string{npmShim}, os.Args[1:]...)...)
	child.Stdin = os.Stdin
	child.Stdout = os.Stdout
	child.Stderr = os.Stderr
	child.Env = os.Environ()
	if dir, derr := os.Getwd(); derr == nil {
		child.Dir = dir
	}

	if err := child.Run(); err != nil {
		if exitErr, ok := err.(*exec.ExitError); ok {
			os.Exit(exitErr.ExitCode())
		}
		fmt.Fprintf(os.Stderr, "codegraph compatibility error: %v\n", err)
		os.Exit(1)
	}
}

func findNpmShim() (string, error) {
	home, herr := os.UserHomeDir()
	if herr != nil || home == "" {
		home = os.Getenv("USERPROFILE")
	}
	if home == "" {
		return "", fmt.Errorf("could not resolve user home directory")
	}

	if override := strings.TrimSpace(os.Getenv("CODEGRAPH_NPM_SHIM")); override != "" {
		if ok, err := verifyShim(override); err == nil && ok {
			return override, nil
		} else if err != nil {
			return "", fmt.Errorf("CODEGRAPH_NPM_SHIM=%s: %w", override, err)
		}
	}

	// First try to discover the real npm global root via `npm root -g`. This
	// respects any custom npm prefix the user has set (e.g. ~/.npm-global/ on
	// Linux, %AppData%\Roaming\npm on Windows) and works on every platform
	// without hardcoding path layouts.
	if npmRoot, err := discoverNpmRoot(); err == nil && npmRoot != "" {
		if shim, ok, _ := resolveShim(filepath.Join(npmRoot, "@aroman22", "codegraph-vba", "npm-shim.js")); ok {
			return shim, nil
		}
	}

	candidates := []string{
		// Windows
		filepath.Join(home, "AppData", "Roaming", "npm", "node_modules", "@aroman22", "codegraph-vba", "npm-shim.js"),
		filepath.Join(home, "AppData", "Local", "pnpm", "global", "5", "node_modules", "@aroman22", "codegraph-vba", "npm-shim.js"),
		filepath.Join(home, "node_modules", "@aroman22", "codegraph-vba", "npm-shim.js"),
		// Linux / macOS
		filepath.Join(home, ".npm-global", "lib", "node_modules", "@aroman22", "codegraph-vba", "npm-shim.js"),
		filepath.Join(home, ".local", "lib", "node_modules", "@aroman22", "codegraph-vba", "npm-shim.js"),
		filepath.Join(home, ".npm", "lib", "node_modules", "@aroman22", "codegraph-vba", "npm-shim.js"),
		"/usr/local/lib/node_modules/@aroman22/codegraph-vba/npm-shim.js",
		"/usr/lib/node_modules/@aroman22/codegraph-vba/npm-shim.js",
	}

	for _, raw := range candidates {
		resolved, ok, err := resolveShim(raw)
		if err != nil {
			continue
		}
		if ok {
			return resolved, nil
		}
	}

	return "", fmt.Errorf("npm-shim.js not found for @aroman22/codegraph-vba. Tried: %s", strings.Join(candidates, "; "))
}

func discoverNpmRoot() (string, error) {
	npm, err := exec.LookPath("npm")
	if err != nil {
		return "", err
	}
	out, err := exec.Command(npm, "root", "-g").Output()
	if err != nil {
		return "", err
	}
	return strings.TrimSpace(string(out)), nil
}

func resolveShim(path string) (string, bool, error) {
	info, err := os.Stat(path)
	if err != nil {
		return "", false, err
	}
	if info.IsDir() {
		return "", false, fmt.Errorf("%s is a directory", path)
	}
	resolved, err := filepath.EvalSymlinks(path)
	if err != nil {
		// EvalSymlinks can fail if the symlink target is missing; fall back to the
		// original path because the launcher still needs to error cleanly.
		return path, true, nil
	}
	if rinfo, serr := os.Stat(resolved); serr != nil || rinfo.IsDir() {
		// Symlink target is broken or a directory; treat as missing.
		return "", false, fmt.Errorf("symlink target %q is unusable", resolved)
	}
	return resolved, true, nil
}

func verifyShim(path string) (bool, error) {
	_, ok, err := resolveShim(path)
	return ok, err
}