


# Load Due Dates from JSON - this will come from manager dashboard
def load_due_dates():
    try:
        with open('due_dates.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

# Load Reference Names from JSON - comes from manager dashboard
def load_reference_names():
    try:
        with open('reference_names.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    
def get_operation_status_enhanced(operations):
    """
    Enhanced function to determine the overall status of operations based on given scenarios.
    """
    statuses = []
    for operation in operations:
        planned_hours = operation['planned_hours']
        actual_hours = operation['actual_hours']

        # Rule 1: Complete if actual_hours >= planned_hours
        if actual_hours >= planned_hours:
            status = 'Complete'

        # Rule 2: In Progress if actual_hours > 0 but not Complete
        elif actual_hours > 0:
            status = 'In Progress'

        # Rule 3: Complete if within 70% of planned hours and subsequent operations are Complete/In Progress
        elif (
            actual_hours >= 0.7 * planned_hours and
            operation.get('subsequent_complete', False)
        ):
            status = 'Complete'

        # Rule 4: Complete if subsequent operations are In Progress and current actual_hours >= 30%
        elif (
            operation.get('subsequent_in_progress', False) and
            actual_hours >= 0.3 * planned_hours
        ):
            status = 'Complete'

        # Rule 5: Not Started otherwise
        else:
            status = 'Not Started'

        statuses.append(status)

    return statuses
    

def calculate_projected_hours(job_operations):
    """
    Calculates projected hours for a job by determining what work remains.
    This function incorporates advanced logic to determine operation statuses and track progress.
    """
    total_actual_hours_projection = 0
    total_remaining_hours = 0
    completed_operations = 0
    total_operations = len(job_operations)

    # Step 1: Sort operations in order
    job_operations.sort(key=lambda op: op.operation_number)

    for idx, op in enumerate(job_operations):
        planned_hours = op.planned_hours or 0
        actual_hours = op.actual_hours or 0
        work_center = op.work_center
        operation_number = op.operation_number
        operation_description = op.task_description if isinstance(op.task_description, str) else ""

        # Skip "Dismantling & Inspection" operations
        if operation_description == "Dismantling & Inspection":
            op.status = 'Ignored'
            continue

        # Step 2: Check subsequent operations
        subsequent_operations = [sub_op for sub_op in job_operations if sub_op.operation_number > operation_number]
        subsequent_complete = all(
            get_operation_status_enhanced([{
                'planned_hours': sub_op.planned_hours,
                'actual_hours': sub_op.actual_hours
            }])[0] in ['Complete', 'In Progress']
            for sub_op in subsequent_operations
        )

        # Step 3: Determine status using enhanced rules
        if work_center == "SR" and planned_hours < 1:
            status = 'Ignored'
        elif work_center in ["ASSEMBLY", "DNI", "WELD"] and actual_hours > 0.8 * planned_hours:
            status = 'Complete'
        elif "Complete" in operation_description and actual_hours > 0.8 * planned_hours:
            status = 'Complete'
        elif planned_hours <= 1 and actual_hours == 0 and subsequent_complete:
            status = 'Complete'
        elif planned_hours < 3 and actual_hours == 0 and subsequent_complete:
            status = 'Complete'
        elif actual_hours > planned_hours:
            status = 'Complete'
        elif planned_hours == 0 and actual_hours == 0 and subsequent_complete:
            status = 'Complete'
        elif planned_hours < 0.5 and subsequent_complete:
            status = 'Complete'
        elif planned_hours < 1 and subsequent_complete:
            status = 'Complete'
        else:
            status = get_operation_status_enhanced([{
                'planned_hours': planned_hours,
                'actual_hours': actual_hours
            }])[0]

        op.status = status  # Assign status to operation

    # Step 4: Handle three consecutive in-progress operations
    in_progress_indices = [i for i, op in enumerate(job_operations) if op.status == 'In Progress']
    for i in range(len(in_progress_indices) - 2):
        if (
            in_progress_indices[i + 1] == in_progress_indices[i] + 1 and
            in_progress_indices[i + 2] == in_progress_indices[i] + 2
        ):
            job_operations[in_progress_indices[i]].status = 'Complete'
            job_operations[in_progress_indices[i + 1]].status = 'Complete'

    # Step 5: Recalculate projection metrics
    for op in job_operations:
        if op.status != 'Ignored':
            total_actual_hours_projection += op.actual_hours
            if op.status not in ['Complete', 'Ignored']:
                total_remaining_hours += max(op.planned_hours - op.actual_hours, 0)
            elif op.status == 'Complete':
                completed_operations += 1

    # Step 6: Calculate projected hours
    projected_hours = total_actual_hours_projection + total_remaining_hours

    return {
        'total_actual_hours_projection': total_actual_hours_projection,
        'total_remaining_hours': total_remaining_hours,
        'projected_hours': projected_hours,
        'completed_operations': completed_operations,
        'total_operations': total_operations
    }


@app.route('/api/jobs')
def get_jobs():
    """Fetch ACTIVE jobs with detailed work orders & operations, including enhanced tracking data."""
    jobs = Job.query.filter_by(active=True).all()
    logging.info(f"📌 Found {len(jobs)} active jobs in database")

    if not jobs:
        return jsonify([])

    # Load external reference data
    due_dates = load_due_dates()
    reference_names = load_reference_names()

    result = []
    today = datetime.now().date()

    di_part_names = {
        "dismantling & inspect",
        "dismantling and inspect",
        "dismantling & inspection",
        "dismantling and inspection",
        "dni"
    }

    ignored_work_orders = {"office use only - do not issue to shop"}

    for job in jobs:
        job_operations = []
        seen_ops_global = set()

        for wo in job.work_orders:
            for op in wo.operations:
                op_key = (op.operation_number, op.task_description, op.work_order_id)
                if op_key not in seen_ops_global:
                    seen_ops_global.add(op_key)
                    job_operations.append(op)

        projected_data = calculate_projected_hours(job_operations)

        total_planned_hours = sum(op.planned_hours or 0 for op in job_operations)
        total_actual_hours = sum(op.actual_hours or 0 for op in job_operations)
        total_remaining_work = sum(op.remaining_work or 0 for op in job_operations)

        completion_percentage = round((total_actual_hours / total_planned_hours) * 100, 1) if total_planned_hours > 0 else 0

        due_date_str = due_dates.get(job.job_number, "To Be Determined")
        due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date() if due_date_str not in ["To Be Determined", None] else None
        days_until_due = (due_date - today).days if due_date else "N/A"

        is_overdue = due_date and days_until_due < 0
        is_at_risk = total_remaining_work > (0.5 * total_planned_hours) and days_until_due != "N/A" and days_until_due <= 5

        num_work_orders = len(job.work_orders)
        num_operations = len(job_operations)

        real_work_orders = [
            wo for wo in job.work_orders
            if not any(
                op.part_name and op.part_name.strip().lower() in ignored_work_orders
                for op in wo.operations
            )
        ]

        is_di_job = all(
            all(
                op.part_name and op.part_name.strip().lower() in di_part_names
                for op in wo.operations
            )
            for wo in real_work_orders
        )

        customer_name = job.customer_name if hasattr(job, 'customer_name') and job.customer_name else "Unknown"

        logging.info(f"   🔹 Job {job.job_number} Classified as D&I: {is_di_job}")

        job_data = {
            'job_number': job.job_number,
            'customer': customer_name,
            'reference_name': reference_names.get(job.job_number, "To Be Determined"),
            'start_date': job.job_start_date.isoformat() if job.job_start_date else None,
            'status': "Completed" if all(op.status == "Complete" for op in job_operations) else "In Progress",
            'due_date': due_date_str,
            'total_planned_hours': total_planned_hours,
            'total_actual_hours': total_actual_hours,
            'projected_hours': projected_data['projected_hours'],
            'remaining_work': total_remaining_work,
            'completion_percentage': completion_percentage,
            'days_until_due': days_until_due,
            'is_overdue': is_overdue,
            'is_at_risk': is_at_risk,
            'num_work_orders': num_work_orders,
            'num_operations': num_operations,
            'is_di_job': is_di_job,
            'work_orders': []
        }

        for wo in job.work_orders:
            work_order_data = {
                'work_order_number': wo.work_order_number,
                'operations': []
            }

            seen_ops = set()
            for op in wo.operations:
                op_key = (op.operation_number, op.task_description)
                if op_key in seen_ops:
                    continue
                seen_ops.add(op_key)

                operation_data = {
                    'operation_number': op.operation_number,
                    'work_center': op.work_center,
                    'part_name': op.part_name,
                    'task_description': op.task_description,
                    'planned_hours': op.planned_hours,
                    'actual_hours': op.actual_hours,
                    'remaining_work': op.remaining_work,
                    'status': op.status,
                    'operation_start_date': op.operation_start_date.isoformat() if op.operation_start_date else None,
                    'scheduled_date': op.scheduled_date.isoformat() if op.scheduled_date else None,
                    'completed_at': op.completed_at.isoformat() if op.completed_at else None
                }
                work_order_data['operations'].append(operation_data)

            job_data['work_orders'].append(work_order_data)

        result.append(job_data)

    logging.info(f"📌 API `/api/jobs` returned {len(result)} jobs")
    return jsonify(result)


#Daily updates code for main dashboard----------------------------------------------------
@app.route('/api/daily_updates', methods=['GET'])
def daily_updates():
    logging.info("🔍 Starting daily update check...")

    latest_worklog_date = db.session.query(func.max(WorkLog.posting_date)).scalar()
    latest_sap_update = db.session.query(func.max(func.date(Operation.updated_at))).scalar()

    if not latest_worklog_date and not latest_sap_update:
        return jsonify({"message": "No data available", "latest_date": None, "updates": {}})

    anchor_date = max(filter(None, [latest_worklog_date, latest_sap_update]))
    logging.info(f"📌 Using {anchor_date} as anchor date")

    # ✅ WorkLogs from that day
    worklogs = WorkLog.query.filter(WorkLog.posting_date == anchor_date).all()
    logging.info(f"📝 Found {len(worklogs)} worklogs")

    worklog_map = defaultdict(list)
    for wl in worklogs:
        key = (wl.job_number, wl.work_order, wl.operation_number)
        worklog_map[key].append({
            "employee_name": wl.employee_name,
            "start_time": wl.start_time.strftime('%H:%M:%S'),
            "end_time": wl.end_time.strftime('%H:%M:%S')
        })

    # ✅ Updated Operations on this day
    updated_ops = Operation.query.filter(func.date(Operation.updated_at) == anchor_date).all()

    # ✅ Completed Operations
    recent_completions = [
        op for op in updated_ops
        if op.status == "Completed" and op.completed_at and op.completed_at.date() == anchor_date
    ]
    logging.info(f"✅ Found {len(recent_completions)} completed operations")

    # ✅ Progressed Operations
    progressed_ops = [
        op for op in updated_ops
        if op.status != "Completed" and op.actual_hours > 0 and op.updated_at.date() == anchor_date
    ]

    # ✅ Build Job Progress Map
    job_progress_map = {}
    for op in progressed_ops:
        job = op.work_order.job if op.work_order and op.work_order.job else None
        if not job:
            continue

        job_key = job.job_number
        job_entry = job_progress_map.setdefault(job_key, {
            "job_number": job.job_number,
            "customer_name": job.customer_name or "Unknown",
            "actual_hours": 0,
            "parts": defaultdict(list)
        })

        job_entry["actual_hours"] += op.actual_hours

        work_key = (job.job_number, op.work_order.work_order_number, op.operation_number)
        employees = worklog_map.get(work_key, [])

        job_entry["parts"][op.part_name or "Unnamed Part"].append({
            "operation_number": op.operation_number,
            "task_description": op.task_description,
            "planned_hours": op.planned_hours,
            "actual_hours": op.actual_hours,
            "status": op.status,
            "updated_at": op.updated_at.isoformat(),
            "employees": employees
        })

    # ✅ Convert parts to list + enrich analytics
    job_progress_response = []
    for job_data in job_progress_map.values():
        total_planned = 0
        total_actual = 0
        completed_credit = 0

        parts_list = []
        for part_name, operations in job_data["parts"].items():
            parts_list.append({
                "part_name": part_name,
                "operations": operations
            })

            for op in operations:
                total_planned += op["planned_hours"]
                total_actual += op["actual_hours"]
                if op["status"] == "Completed":
                    completed_credit += op["planned_hours"]
                elif op["status"] == "In Progress":
                    completed_credit += min(op["actual_hours"], op["planned_hours"])

        job_data["parts"] = parts_list
        job_data["analytics"] = {
            "total_planned": round(total_planned, 2),
            "total_actual": round(total_actual, 2),
            "remaining_work": round(max(total_planned - total_actual, 0), 2),
            "completion_percent": round((completed_credit / total_planned) * 100, 2) if total_planned else 0
        }

        job_progress_response.append(job_data)

    # ✅ Missing Clock-ins
    today = date.today()
    recent_weekdays = [today - timedelta(days=i) for i in range(1, 8) if (today - timedelta(days=i)).weekday() < 5]
    all_employees = [e[0] for e in db.session.query(WorkLog.employee_name).distinct().all()]
    missing_days = []
    for emp in all_employees:
        worked_dates = {
            d[0] for d in db.session.query(WorkLog.posting_date).filter_by(employee_name=emp).distinct().all()
        }
        for d in recent_weekdays:
            if d not in worked_dates:
                missing_days.append({"employee": emp, "missing_date": d.isoformat()})

    return jsonify({
        "latest_date": anchor_date.isoformat(),
        "worklogs": [log.to_dict() for log in worklogs],
        "recent_completions": [  # Keep this basic for now
            {
                "job_number": op.work_order.job.job_number if op.work_order and op.work_order.job else None,
                "work_order": op.work_order.work_order_number if op.work_order else None,
                "operation_number": op.operation_number,
                "work_center": op.work_center,
                "part_name": op.part_name,
                "status": op.status,
                "actual_hours": op.actual_hours,
                "planned_hours": op.planned_hours,
                "completed_at": op.completed_at.isoformat() if op.completed_at else None,
                "updated_at": op.updated_at.isoformat() if op.updated_at else None
            } for op in recent_completions
        ],
        "job_progress": job_progress_response,
        "missing_clockin_days": missing_days
    })

#End Daily Updates API code------------------------------------------------------------------------------------



# Start Manager Dashboard--------------------------------------------


@app.route("/manager")
def manager_dashboard():
    return render_template("manager_dashboard.html")


@app.route("/api/manager/active_jobs")
def manager_active_jobs():
    active_summary, di_summary, totals = get_active_job_summary_from_db()
    return jsonify({
        "active_summary": active_summary,
        "di_summary": di_summary,
        "totals": totals
    })

def get_active_job_summary_from_db():
    """
    Returns active and D&I job summaries using database models including PO integration.
    """
    reference_names = load_reference_names()
    due_dates = load_due_dates()
    order_values = get_order_values()

    active_summary = []
    di_summary = []
    total_values = {
        "total_jobs": 0,
        "total_planned_hours": 0,
        "total_actual_hours": 0,
        "total_projected_hours": 0,
        "total_planned_cost": 0,
        "total_actual_cost": 0,
        "total_projected_cost": 0,
        "total_order_value": 0,
        "total_profit_value": 0,
        "average_profit_margin": 0,
    }
    profit_margins = []

    jobs = Job.query.options(
        joinedload(Job.work_orders).joinedload(WorkOrder.operations)
    ).filter_by(active=True).all()

    for job in jobs:
        job_number = job.job_number
        all_operations = [op for wo in job.work_orders for op in wo.operations]

        if not all_operations:
            continue

        part_names = set(op.part_name for op in all_operations if op.part_name)
        is_di_job = all(name == "Dismantling & Inspection" or sum(op.planned_hours for op in all_operations if op.part_name == name) == 0 for name in part_names)

        # 🧮 Operation Hours
        total_planned_hours = sum(op.planned_hours for op in all_operations)
        total_actual_hours = sum(op.actual_hours for op in all_operations)
        remaining_hours = sum(op.remaining_work for op in all_operations)
        projected_hours = total_actual_hours + remaining_hours

        # 💲 Labor Costs
        total_planned_labor_cost = sum(
            calculate_cost(op.planned_hours, op.part_name, op.work_center, op.task_description)
            for op in all_operations
        )
        total_actual_labor_cost = sum(
            calculate_cost(op.actual_hours, op.part_name, op.work_center, op.task_description)
            for op in all_operations
        )

        # 📦 Purchase Orders (from DB now)
        po_records = PurchaseOrder.query.filter_by(job_number=job_number).all()
        total_goods_cost = sum(po.net_price * po.order_quantity for po in po_records)
        still_to_be_delivered_value = sum(po.pending_value for po in po_records)
        cost_goods_received = total_goods_cost - still_to_be_delivered_value

        # 🛡️ Warranty/Overhead
        order_value = order_values.get(job_number)
        warranty_cost = order_value * 0.015 if order_value else 0

        # 📊 Cost Summaries
        total_planned_cost = total_planned_labor_cost + total_goods_cost + warranty_cost
        total_actual_cost = total_actual_labor_cost + cost_goods_received + warranty_cost
        projected_cost = (
            calculate_cost(total_actual_hours) +
            calculate_cost(remaining_hours) +
            total_goods_cost
        )

        # 📈 Profit
        if order_value:
            profit_value = order_value - total_actual_cost
            profit_margin = (profit_value / order_value) * 100
        else:
            profit_value = None
            profit_margin = None

        # 📅 Metadata
        reference_name = reference_names.get(job_number, "")
        due_date_str = due_dates.get(job_number)
        try:
            due_date = datetime.strptime(due_date_str, "%Y-%m-%d") if due_date_str else None
        except ValueError:
            due_date = None
        due_date_formatted = due_date.strftime('%Y-%m-%d') if due_date else "No Due Date Assigned"

        job_data = {
            'job_number': job_number,
            'customer': job.customer_name or "Unknown Customer",
            'total_planned_hours': format_number(total_planned_hours),
            'total_actual_hours': format_number(total_actual_hours),
            'projected_hours': format_number(projected_hours),
            'total_planned_cost': format_number(total_planned_cost),
            'total_actual_cost': format_number(total_actual_cost),
            'projected_cost': format_number(projected_cost),
            'order_value': format_number(order_value) if order_value else 'N/A',
            'profit_value': format_number(profit_value) if profit_value else 'N/A',
            'profit_margin': format_number(profit_margin) if profit_margin else 'N/A',
            'reference_name': reference_name,
            'due_date': due_date,
            'due_date_formatted': due_date_formatted,
        }

        if is_di_job:
            di_summary.append(job_data)
        else:
            active_summary.append(job_data)

        # Totals
        total_values["total_jobs"] += 1
        total_values["total_planned_hours"] += total_planned_hours
        total_values["total_actual_hours"] += total_actual_hours
        total_values["total_projected_hours"] += projected_hours
        total_values["total_planned_cost"] += total_planned_cost
        total_values["total_actual_cost"] += total_actual_cost
        total_values["total_projected_cost"] += projected_cost
        if order_value:
            total_values["total_order_value"] += order_value
        if profit_value:
            total_values["total_profit_value"] += profit_value
        if profit_margin:
            profit_margins.append(profit_margin)

    if profit_margins:
        total_values["average_profit_margin"] = sum(profit_margins) / len(profit_margins)

    total_values = {
        key: format_number(value) if key != "total_jobs" else value
        for key, value in total_values.items()
    }

    active_summary.sort(key=lambda x: x['due_date'] or datetime.max)
    di_summary.sort(key=lambda x: x['due_date'] or datetime.max)

    return active_summary, di_summary, total_values


def get_order_values():
    """
    Load order values from order_values.json in the same directory as this script.
    """
    try:
        file_path = os.path.join(os.path.dirname(__file__), 'order_values.json')
        with open(file_path, 'r') as file:
            order_values = json.load(file)
        return {str(k): v for k, v in order_values.items()}
    except FileNotFoundError:
        print("Error: order_values.json file not found.")
        return {}
    except json.JSONDecodeError:
        print("Error: order_values.json is not valid JSON.")
        return {}
    except Exception as e:
        print(f"Error loading order values: {e}")
        return {}

def calculate_cost(hours, description=None, work_center=None, task_description=None):
    """
    Calculate labor cost with reduced burden rate for engineering/admin/RC-type tasks.
    """
    default_burden_rate = 199
    reduced_burden_rate = 10

    if (
        description in ['RC', 'Engineering', 'Admin', 'RC / Engineering / Admin.'] or 
        work_center == 'REP ENG' or 
        task_description == 'Engineering Time'
    ):
        burden_rate = reduced_burden_rate
    else:
        burden_rate = default_burden_rate

    return hours * burden_rate

def format_number(value):
    """
    Format numbers for display, fallback to string if error occurs.
    """
    if value is None:
        return 'N/A'
    try:
        return "{:,.2f}".format(value)
    except:
        return str(value)
