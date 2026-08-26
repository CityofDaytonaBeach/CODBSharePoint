// ============================================================================
// Security Scanner - Analyzes code for security issues
// ============================================================================

import type {
  CODBIR,
  SecurityReport,
  SecurityFinding,
  SecretFinding,
  PermissionAudit
} from '../types/index.js';
import { GRAPH_PERMISSIONS } from '../types/index.js';

export class SecurityScanner {

  async scan(ir: CODBIR, sourceFiles?: Map<string, string>): Promise<SecurityReport> {
    const findings: SecurityFinding[] = [];
    const secrets: SecretFinding[] = [];
    const externalUrls: string[] = [];
    const permissions: PermissionAudit[] = [];
    const recommendations: string[] = [];

    // Scan source files if provided
    if (sourceFiles) {
      this.scanSourceFiles(sourceFiles, findings, secrets, externalUrls);
    }

    // Scan IR for security issues
    this.scanIR(ir, findings, externalUrls);

    // Audit permissions
    this.auditPermissions(ir, permissions);

    // Generate recommendations
    this.generateRecommendations(ir, findings, permissions, recommendations);

    // Calculate score
    const score = this.calculateScore(findings, secrets);

    return {
      passed: findings.filter(f => f.severity === 'critical' || f.severity === 'high').length === 0,
      score,
      findings,
      secrets,
      externalUrls: [...new Set(externalUrls)],
      permissions,
      recommendations
    };
  }

  private scanSourceFiles(
    files: Map<string, string>,
    findings: SecurityFinding[],
    secrets: SecretFinding[],
    externalUrls: string[]
  ): void {
    for (const [path, content] of files) {
      // Check for secrets
      this.findSecrets(path, content, secrets);

      // Check for dangerous patterns
      this.findDangerousPatterns(path, content, findings);

      // Extract URLs
      this.extractUrls(path, content, externalUrls);
    }
  }

  private findSecrets(file: string, content: string, secrets: SecretFinding[]): void {
    const patterns: Array<{ type: string; regex: RegExp }> = [
      { type: 'API Key', regex: /(?:api[_-]?key|apikey)\s*[:=]\s*['"]([a-zA-Z0-9_-]{20,})['"]/gi },
      { type: 'Secret', regex: /(?:secret|password|passwd)\s*[:=]\s*['"]([^'"]{8,})['"]/gi },
      { type: 'Token', regex: /(?:token|access_token|auth_token)\s*[:=]\s*['"]([a-zA-Z0-9._-]{20,})['"]/gi },
      { type: 'Private Key', regex: /-----BEGIN\s+(?:RSA\s+)?PRIVATE\s+KEY-----/gi },
      { type: 'Connection String', regex: /(?:Server|Data Source|Host)\s*=\s*[^;]+/gi },
      { type: 'AWS Key', regex: /(?:AKIA|ABIA|ACCA|ASIA)[A-Z0-9]{16}/g },
      { type: 'GitHub Token', regex: /ghp_[a-zA-Z0-9]{36}/g },
      { type: 'Azure Key', regex: /(?:AccountKey|SharedAccessSignature|DefaultEndpointsProtocol)[^'"]{20,}/gi }
    ];

    const lines = content.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      for (const { type, regex } of patterns) {
        let match;
        while ((match = regex.exec(line)) !== null) {
          const value = match[1] || match[0];
          const masked = value.substring(0, 4) + '****' + value.substring(value.length - 4);
          secrets.push({
            type,
            file,
            line: i + 1,
            masked
          });
        }
      }
    }
  }

  private findDangerousPatterns(file: string, content: string, findings: SecurityFinding[]): void {
    const patterns: Array<{
      severity: 'critical' | 'high' | 'medium' | 'low';
      category: string;
      rule: string;
      message: string;
      regex: RegExp;
    }> = [
      {
        severity: 'critical',
        category: 'injection',
        rule: 'eval-usage',
        message: 'eval() detected - potential code injection',
        regex: /\beval\s*\(/g
      },
      {
        severity: 'high',
        category: 'injection',
        rule: 'innerhtml-usage',
        message: 'innerHTML detected - potential XSS vulnerability',
        regex: /\.innerHTML\s*=/g
      },
      {
        severity: 'high',
        category: 'injection',
        rule: 'document-write',
        message: 'document.write() detected - potential XSS vulnerability',
        regex: /document\.write\s*\(/g
      },
      {
        severity: 'medium',
        category: 'injection',
        rule: 'dangerouslysetinnerhtml',
        message: 'dangerouslySetInnerHTML detected - review for XSS',
        regex: /dangerouslySetInnerHTML/g
      },
      {
        severity: 'medium',
        category: 'network',
        rule: 'http-endpoint',
        message: 'HTTP endpoint detected (not HTTPS)',
        regex: /http:\/\/(?!localhost|127\.0\.0\.1)/gi
      },
      {
        severity: 'medium',
        category: 'secrets',
        rule: 'hardcoded-url',
        message: 'Hardcoded external URL detected',
        regex: /https?:\/\/[^\s'"`]+/g
      },
      {
        severity: 'low',
        category: 'best-practice',
        rule: 'console-log',
        message: 'console.log() detected - remove for production',
        regex: /console\.(log|debug|info)\s*\(/g
      },
      {
        severity: 'low',
        category: 'best-practice',
        rule: 'debugger',
        message: 'debugger statement detected',
        regex: /\bdebugger\b/g
      }
    ];

    const lines = content.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      for (const { severity, category, rule, message, regex } of patterns) {
        if (regex.test(line)) {
          findings.push({
            severity,
            category,
            rule,
            message,
            file,
            line: i + 1,
            evidence: line.trim().substring(0, 100)
          });
        }
      }
    }
  }

  private extractUrls(file: string, content: string, urls: string[]): void {
    const urlRegex = /https?:\/\/[^\s'"`<>]+/g;
    let match;
    while ((match = urlRegex.exec(content)) !== null) {
      urls.push(match[0]);
    }
  }

  private scanIR(ir: CODBIR, findings: SecurityFinding[], externalUrls: string[]): void {
    // Check solution configuration
    if (ir.solution.isDomainIsolated) {
      findings.push({
        severity: 'info',
        category: 'deployment',
        rule: 'domain-isolated',
        message: 'Solution is domain isolated - limited SharePoint API access',
        fix: 'Ensure this is intentional for the solution requirements'
      });
    }

    // Check for excessive permissions
    const adminPermissions = ir.graph.filter(p => p.requiresAdminApproval);
    if (adminPermissions.length > 5) {
      findings.push({
        severity: 'medium',
        category: 'permissions',
        rule: 'excessive-permissions',
        message: `Solution requires ${adminPermissions.length} admin-approved Graph permissions`,
        fix: 'Review if all permissions are necessary'
      });
    }

    // Check for sensitive permission combinations
    const sensitiveScopes = ['User.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All'];
    const requestedSensitive = ir.graph.filter(p => sensitiveScopes.includes(p.scope));
    if (requestedSensitive.length > 0) {
      findings.push({
        severity: 'medium',
        category: 'permissions',
        rule: 'sensitive-permissions',
        message: `Solution requests sensitive permissions: ${requestedSensitive.map(p => p.scope).join(', ')}`,
        fix: 'Ensure these permissions are required and document the justification'
      });
    }
  }

  private auditPermissions(ir: CODBIR, permissions: PermissionAudit[]): void {
    // All declared permissions
    for (const perm of ir.graph) {
      const knownPerm = GRAPH_PERMISSIONS[perm.scope as keyof typeof GRAPH_PERMISSIONS];
      permissions.push({
        permission: perm.scope,
        required: true,
        declared: true,
        actuallyUsed: true, // Would need code analysis to determine
        risk: knownPerm?.requiresAdminApproval ? 'high' : 'low'
      });
    }
  }

  private generateRecommendations(
    ir: CODBIR,
    findings: SecurityFinding[],
    permissions: PermissionAudit[],
    recommendations: string[]
  ): void {
    if (findings.some(f => f.severity === 'critical')) {
      recommendations.push('Address critical security findings before deployment');
    }

    if (findings.some(f => f.rule === 'eval-usage')) {
      recommendations.push('Replace eval() with safer alternatives like JSON.parse()');
    }

    if (findings.some(f => f.rule === 'innerhtml-usage')) {
      recommendations.push('Use React\'s JSX or DOM API instead of innerHTML');
    }

    const highRiskPerms = permissions.filter(p => p.risk === 'high');
    if (highRiskPerms.length > 0) {
      recommendations.push(`Review high-risk permissions: ${highRiskPerms.map(p => p.permission).join(', ')}`);
    }

    if (findings.some(f => f.rule === 'http-endpoint')) {
      recommendations.push('Use HTTPS for all external communications');
    }

    if (findings.some(f => f.rule === 'console-log')) {
      recommendations.push('Remove console.log() statements before production deployment');
    }
  }

  private calculateScore(findings: SecurityFinding[], secrets: SecretFinding[]): number {
    let score = 100;

    // Deduct for findings
    for (const finding of findings) {
      switch (finding.severity) {
        case 'critical': score -= 25; break;
        case 'high': score -= 15; break;
        case 'medium': score -= 8; break;
        case 'low': score -= 3; break;
      }
    }

    // Deduct for secrets
    score -= secrets.length * 20;

    return Math.max(0, score);
  }
}
