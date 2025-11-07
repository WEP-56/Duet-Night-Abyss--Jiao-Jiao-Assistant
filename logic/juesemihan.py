import os
import time
import win32gui


def _sleep_interruptible(app, seconds, step=0.05):
    waited = 0.0
    while waited < seconds and app.running and not app.stop_event.is_set():
        time.sleep(step)
        waited += step


def _select_mihan(app):
    """滚轮搜索并选择用户在GUI中选定的角色密函图片。返回True表示成功定位到目标。"""
    target_tpl = app.get_selected_juese_mihan_path()
    if not target_tpl or not os.path.isfile(target_tpl):
        app._log('未选择角色密函或模板不存在，跳过选择。')
        return False

    app._log('🔍 正在选择角色密函')
    # 将鼠标移到"不使用"按钮中心，便于滚轮居中
    bushiyong = os.path.join(app.control_dir, 'bushiyong.png')
    m = app.detect_template_abs_scales(bushiyong, scales=[1.1, 1.05, 1.0, 0.95, 0.9], threshold=0.80)
    wheel_pos = None
    if m:
        target, tx, ty = app.center_to_client_and_target(m['center'])
        wheel_pos = (tx, ty)
        app._log(f"🖱️ 移动到不使用中心以便滚轮: client({tx},{ty})")
        # 可选轻触，避免阻塞
        app._try_wait_and_click('bushiyong.png', 'bushiyong', timeout=0.3)
        _sleep_interruptible(app, 0.1)
    else:
        # fallback to client center (avoid using external helpers)
        try:
            left, top, right, bottom = win32gui.GetClientRect(app.selected_hwnd)
            cx = max(0, (right - left) // 2)
            cy = max(0, (bottom - top) // 2)
            wheel_pos = (cx, cy)
            app._log(f"⚠️ 未识别到不使用，使用窗口中心滚轮: client{wheel_pos}")
        except Exception:
            wheel_pos = None

    # 先向下滚10次，每次滚后识别一次；命中则点击该密函并确认
    for _ in range(10):
        if app.send_mouse_wheel(delta=-120, count=1, client_pos=wheel_pos):
            _sleep_interruptible(app, 0.1)
        m_hit = app.detect_template_abs_scales(target_tpl, scales=[1.1, 1.05, 1.0, 0.95, 0.9], threshold=0.80)
        if m_hit:
            app._log('✅ 已定位到所选角色密函，执行点击与确认')
            app.click_match_abs(target_tpl, 'mihan_target', threshold=0.80, scales=[1.1, 1.05, 1.0, 0.95, 0.9])
            _sleep_interruptible(app, 0.2)
            app.click_match_abs(os.path.join(app.control_dir, 'querenxuanze.png'), 'querenxuanze', threshold=0.80, scales=[1.0, 0.95, 0.9, 1.05, 1.1])
            return True
    # 再向上滚10次
    for _ in range(10):
        if app.send_mouse_wheel(delta=120, count=1, client_pos=wheel_pos):
            _sleep_interruptible(app, 0.1)
        m_hit = app.detect_template_abs_scales(target_tpl, scales=[1.1, 1.05, 1.0, 0.95, 0.9], threshold=0.80)
        if m_hit:
            app._log('✅ 已定位到所选角色密函，执行点击与确认')
            app.click_match_abs(target_tpl, 'mihan_target', threshold=0.80, scales=[1.1, 1.05, 1.0, 0.95, 0.9])
            _sleep_interruptible(app, 0.2)
            app.click_match_abs(os.path.join(app.control_dir, 'querenxuanze.png'), 'querenxuanze', threshold=0.80, scales=[1.0, 0.95, 0.9, 1.05, 1.1])
            return True
    app._log('❌ 未定位到所选角色密函')
    return False


def _reward_select(app):
    """奖励选择策略：first > second(避开词缀second) > third(优先角色经验)。
    简化实现：当前模板匹配返回单一位置，无法严格在多实例中二选一，先实现优先级与奖励偏好。
    """
    base = os.path.join(app.control_dir, '奖励选择png')
    def _abs(name):
        return os.path.join(base, f'{name}.png')

    # 1) first
    m = app.detect_template_abs(_abs('first'))
    if m:
        app._log('🎁 奖励选择: 发现 first，点击')
        app.click_match_abs(_abs('first'), 'first')
        return True

    # 2) second: 尽量避开 cishi-second
    sec = app.detect_template_abs(_abs('second'))
    cishi = app.detect_template_abs(_abs('cishi-second'))
    if sec:
        if not cishi:
            app._log('🎁 奖励选择: 发现 second（无词缀），点击')
            app.click_match_abs(_abs('second'), 'second')
            return True
        else:
            app._log('🎁 奖励选择: 发现 second 与 cishi-second，尝试选择非词缀项（简化为点击second）')
            app.click_match_abs(_abs('second'), 'second')
            return True

    # 3) third: 角色经验优先，其次碎片
    third = app.detect_template_abs(_abs('third'))
    if third:
        jy = app.detect_template_abs(_abs('juesejingyan-third'))
        sp = app.detect_template_abs(_abs('suipian-third'))
        if jy:
            app._log('🎁 奖励选择: 发现 third 且角色经验奖励，点击 third')
            app.click_match_abs(_abs('third'), 'third')
            return True
        if sp:
            app._log('🎁 奖励选择: 发现 third 且碎片奖励，点击 third')
            app.click_match_abs(_abs('third'), 'third')
            return True
        # 若仅有third而未识别奖励模板，仍点击
        app._log('🎁 奖励选择: 发现 third（未识别到奖励细分），点击 third')
        app.click_match_abs(_abs('third'), 'third')
        return True

    app._log('🎁 奖励选择: 未识别到可用选项')
    return False


def run(app):
    # Step0: 进入密函界面：点击"选择密函"，等待出现"不使用"作为成功进入标志
    xuanzemihan = os.path.join(app.control_dir, 'xuanzemihan.png')
    app._log(f"🖼️ 选择密函模板: {xuanzemihan} 存在={os.path.isfile(xuanzemihan)}")
    if not app.click_match_abs(xuanzemihan, 'xuanzemihan', threshold=0.80, scales=[1.0, 0.95, 0.9, 1.05, 1.1]):
        app._log('❌ 未识别到 选择密函 按钮')
        return
    _sleep_interruptible(app, 0.5)

    # 显式等待"不使用"出现
    by_path = os.path.join(app.control_dir, 'bushiyong.png')
    app._log('⏳ 等待不使用出现以开始滚轮搜索')
    wait_deadline = time.time() + 5.0
    while app.running and not app.stop_event.is_set() and time.time() < wait_deadline:
        if app.detect_template_abs_scales(by_path, scales=[1.1, 1.05, 1.0, 0.95, 0.9], threshold=0.80):
            app._log('✅ 检测到不使用，开始滚轮搜索密函')
            break
        _sleep_interruptible(app, 0.2)

    # 进入密函选择程序（滚轮搜索并点击选中+确认）
    _select_mihan(app)

    # 点击用户所选密函
    sel = app.get_selected_juese_mihan_path()
    if sel and os.path.isfile(sel):
        app.click_match_abs(sel, os.path.basename(sel))
    else:
        app._log('⚠️ 未选择角色密函，跳过点击所选密函')

    # 确认选择
    if not app._wait_and_click('querenxuanze.png', 'querenxuanze'):
        return

    # 等待进入地图标志
    if not app._wait_detect('likai.png', 'likai'):
        return
    _sleep_interruptible(app, float(app.post_likai_delay))

    # 地图识别
    map_name = app._recognize_map_name()
    if not map_name:
        app._log('🗺️ 地图识别失败，重试一次...')
        _sleep_interruptible(app, 0.3)
        map_name = app._recognize_map_name()
    steps = None
    exec_name = None
    if map_name:
        steps = app._load_actions(map_name)
        exec_name = map_name
    else:
        if not app.fail_fallback_random:
            app._log('🚫 地图仍未识别，且未开启随机脚本策略，停止。')
            app.running = False
            return
        try:
            files = [f for f in os.listdir(app.json_dir) if f.lower().endswith('.json')]
        except Exception:
            files = []
        if not files:
            app._log('📂 无可用脚本可供随机选择，停止。')
            app.running = False
            return
        pick = files[0]
        exec_name = os.path.splitext(pick)[0]
        app._log(f"🎲 地图仍未识别，随机选择脚本: {pick}")
        steps = app._load_actions(exec_name)
    if not steps:
        app._log('📋 未加载到动作步骤，停止。')
        app.running = False
        return

    # 执行动作
    app._log(f"🎮 开始执行 {exec_name} 的移动脚本，共 {len(steps)} 步")
    app.play_actions(app.selected_hwnd, steps, app._log, app.stop_event)
    app._log('🏃 移动操作结束，等待 querenxuanze')

    # 奖励选择
    if not app._wait_and_click('querenxuanze.png', 'querenxuanze'):
        return
    _sleep_interruptible(app, 0.2)
    _reward_select(app)
