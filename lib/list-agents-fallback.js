/**
 * `list_agents` 兼容层。
 *
 * dsh 0.2.1 把 `list_agents` 放在 `@deepseek-ai/dsh-tool-subagent-control/list-agents`
 * 这个独立子入口里；dsh 0.2.2 删掉了该子入口，把工具并进包的主入口。
 * 预设如果在 0.2.2 上继续挂那个子入口，整条预设会被判 broken、桌面端显示「加载失败」。
 *
 * 所以这里自己兜一个：主入口已经注册了同名工具就直接退出，否则补注册。
 * 不 import 任何核心包，避免解析路径随内核版本变化。
 */

export const name = 'ppt-list-agents-fallback'

export const inject = ['tools', 'subagents', 'agents']

function statusOf(agents, id) {
  return agents.get(id)?.status === 'running' ? 'running' : 'inactive'
}

function project(agents, entry, position) {
  const at = position === undefined ? {} : { parent: position.parentId, depth: position.depth }
  if ('kind' in entry && entry.kind === 'diagnostic') {
    return { kind: 'diagnostic', id: entry.id, reason: entry.reason, ...at }
  }
  // 一次性子代理没法用 send_message 续接，模型不会选它，所以不暴露。
  if (entry.mode !== 'continuable') return undefined
  return { kind: 'child', id: entry.id, label: entry.label, status: statusOf(agents, entry.id), ...at }
}

function buildTool(ctx) {
  return {
  name: 'list_agents',
  description:
    'List subagents you started, with their ids, labels, and status. '
    + 'running means it is working; inactive means it is not currently working. '
    + 'You will be notified when a subagent finishes; there is no need to keep checking its status. '
    + 'Use send_message to continue the conversation.',
  parameters: {
    scope: {
      type: 'string',
      enum: ['children', 'descendants'],
      description:
        'children (default) lists direct children, which accept send_message in any status. '
        + "descendants lists the whole tree below you with each entry's parent session id and depth; "
        + 'entries deeper than 1 accept only interrupt_agent.',
    },
  },
  output: {
    schema: {
      type: 'array',
      items: {
        oneOf: [
          {
            type: 'object',
            additionalProperties: false,
            required: ['kind', 'id', 'label', 'status'],
            properties: {
              kind: { type: 'string', enum: ['child'] },
              id: { type: 'string' },
              label: { type: 'string' },
              status: { type: 'string', enum: ['running', 'inactive'] },
              parent: { type: 'string' },
              depth: { type: 'number' },
            },
          },
          {
            type: 'object',
            additionalProperties: false,
            required: ['kind', 'id', 'reason'],
            properties: {
              kind: { type: 'string', enum: ['diagnostic'] },
              id: { type: 'string' },
              reason: { type: 'string', enum: ['corrupt', 'unsupported', 'unavailable'] },
              parent: { type: 'string' },
              depth: { type: 'number' },
            },
          },
        ],
      },
    },
    render: (args, entries) => {
      const scope = args?.scope ?? 'children'
      return [{
        type: 'text',
        text: entries.length === 0
          ? '(no subagents)'
          : entries.map((entry) => {
            const at = scope === 'descendants'
              ? ` parent=${String(entry.parent)} depth=${String(entry.depth)}`
              : ''
            return entry.kind === 'child'
              ? `${entry.id} [${entry.status}]${at} — ${entry.label}`
              : `${entry.id} [diagnostic: ${entry.reason}]${at}`
          }).join('\n'),
      }]
    },
  },
  async execute(args, exec) {
    const parent = exec.agent
    if (!parent) throw new Error('list_agents requires a calling agent (exec.agent was undefined)')
    const scope = args?.scope ?? 'children'
    if (scope === 'descendants') {
      const entries = await ctx.subagents.listDescendants(parent.id, exec.signal)
      return entries.map((entry) => project(ctx.agents, entry, entry)).filter((e) => e !== undefined)
    }
    const entries = await ctx.subagents.listChildren(parent.id, exec.signal)
    return entries.map((entry) => project(ctx.agents, entry)).filter((e) => e !== undefined)
  },
  }
}

export function apply(ctx) {
  // 0.2.2 起主入口已经自带同名工具；这里延后一拍再判断，让主入口先注册，
  // 避免同一 scope 里重复注册把预设打成 broken。0.2.1 没有主入口的同名工具，
  // 这一拍之后补上即可。
  const timer = setTimeout(() => {
    try {
      if (ctx.tools.get?.('list_agents')) return
    } catch {
      /* 查询失败就当没有，交给注册时的 try/catch 兜底 */
    }
    try {
      ctx.tools.register(buildTool(ctx))
    } catch {
      /* 已经被别的入口注册了，让位 */
    }
  }, 50)
  timer.unref?.()
}
