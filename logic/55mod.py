import time


def run(app):
    """Night航55 模式主循环。依赖于 app 中已实现的工具方法和属性。

    Sequence per round:
    - optional click querenxuanze once
    - click kaishitiaozhan to enter
    - wait likai, delay, recognize map, load actions, play
    - wait and click zaicijinixng
    - click kaishitiaozhan to next round
    """
    # 启动时：尽量点一次“确认选择”（可选，超时不致停）
    app._try_wait_and_click('querenxuanze.png', 'querenxuanze', timeout=3.0)
    # 启动时：点一次“开始挑战”进入地图
    if not app._wait_and_click('kaishitiaozhan.png', 'kaishitiaozhan'):
        return

    app.loops_done = 0
    while app.running:
        # 1) 等待进入地图标志
        if not app._wait_detect('likai.png', 'likai'):
            break
        # 延迟，等待场景稳定
        waited = 0.0
        step = 0.05
        extra_delay = float(app.post_likai_delay)
        while waited < extra_delay and app.running and not app.stop_event.is_set():
            time.sleep(step)
            waited += step

        # 2) 识图，失败重试一次；仍失败按设置随机
        map_name = app._recognize_map_name()
        if not map_name:
            app._log('地图识别失败，重试一次...')
            waited = 0.0
            step = 0.05
            retry_delay = 0.3
            while waited < retry_delay and app.running and not app.stop_event.is_set():
                time.sleep(step)
                waited += step
            map_name = app._recognize_map_name()
        if map_name:
            steps = app._load_actions(map_name)
            exec_name = map_name
        else:
            if not app.fail_fallback_random:
                app._log('地图仍未识别，且未开启随机脚本策略，停止。')
                app.running = False
                break
            else:
                import os, random
                try:
                    files = [f for f in os.listdir(app.json_dir) if f.lower().endswith('.json')]
                except Exception:
                    files = []
                if not files:
                    app._log('无可用脚本可供随机选择，停止。')
                    app.running = False
                    break
                pick = random.choice(files)
                pick_name = os.path.splitext(pick)[0]
                app._log(f"地图仍未识别，随机选择脚本: {pick}")
                steps = app._load_actions(pick_name)
                exec_name = pick_name
        if not steps:
            app._log('未加载到动作步骤，停止。')
            app.running = False
            break
        app._log(f"开始执行 {exec_name} 的移动脚本，共 {len(steps)} 步")
        app.play_actions(app.selected_hwnd, steps, app._log, app.stop_event)
        app._log("移动操作结束，等待 zaicijinixng_button")

        # 3) 等待 战斗结束 图标 并点击
        if not app._wait_and_click('zaicijinixng.png', 'zaicijinixng'):
            break

        # 4) 点击 开始挑战_button，开始下一轮
        if not app._wait_and_click('kaishitiaozhan.png', 'kaishitiaozhan'):
            break

        # 循环次数限制
        app.loops_done += 1
        if app.max_loops and app.loops_done >= app.max_loops:
            app._log(f"已完成设定的循环次数 {app.max_loops}，停止运行。")
            app.running = False
            break
