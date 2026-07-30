import re
import sys

html_path = 'G:/My Drive/Dữ liệu nhân sự/.agents/skills/Ops_weekly_output_report/template/weekly_output_dashboard.html'
with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

def replace_brace_obj(text, marker, new_val):
    rm_idx = text.find(marker)
    if rm_idx == -1: return text
    obj_start = rm_idx + len(marker)
    depth, i, in_str, esc = 0, obj_start, False, False
    while i < len(text):
        c = text[i]
        if esc:                    esc = False
        elif c == '\\\\' and in_str: esc = True
        elif c == '"' and not esc: in_str = not in_str
        elif not in_str:
            if c == '{': depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0: break
        i += 1
    old_obj = text[rm_idx: i + 2]
    new_obj = marker + new_val + ';'
    return text.replace(old_obj, new_obj, 1)

# Replace RAW in ganttSource
gantt_start = html.find('<script type="text/plain" id="ganttSource">')
gantt_end   = html.find('</script>', gantt_start)
gantt_content = html[gantt_start:gantt_end]
new_gantt = replace_brace_obj(gantt_content, 'const RAW = ', 'parent.window.GANTT_RAW || {}')
html = html[:gantt_start] + new_gantt + html[gantt_end:]

# Replace CUSTOMER_TARGETS
html = replace_brace_obj(html, 'const CUSTOMER_TARGETS = ', 'window.CUSTOMER_TARGETS || {}')

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print('Brace obj patched successfully.')
