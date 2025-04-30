from flask import Blueprint

# ✅ Define Blueprint for logistics dashboard
main_bp = Blueprint('main', __name__)


import os
import pandas as pd
from flask import render_template, request, jsonify, Response, stream_with_context, current_app, send_file
from app import app, db
from models import db, Job, WorkOrder, Operation, JobHistory, Note, PurchaseOrder, Shipment, Notification, WorkCenter, WorkLog, NCRTracker
from utils import process_sapdata, calculate_forecast, process_worklog
from datetime import datetime, timedelta, date
from werkzeug.utils import secure_filename
import logging
from flask import jsonify
import json
from concurrent.futures import ThreadPoolExecutor
from sqlalchemy.sql import text
import requests
from sqlalchemy.orm import joinedload
from collections import Counter # For Purchase order page for shop lead
import uuid
import shutil
from excel_processor import process_purchase_orders
import os
from sqlalchemy import func, and_ # Used for daily updates section
from collections import defaultdict #Used for daily updates to show worklog and job progress
import json #for order values to be loaded into manager dashboard
import re # Regex for finding task descriptions with similar descriptions

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


# Configure thread pool for background tasks
executor = ThreadPoolExecutor(max_workers=3)

#Start Logisitics dashboard related code---------------------------------------


# Configure upload folder
UPLOAD_FOLDER = 'static/uploads'
SHIPMENT_FILES_FOLDER = os.path.join(UPLOAD_FOLDER, 'shipments')
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf', 'jpg', 'jpeg', 'png', 'gif'}

# Create upload directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SHIPMENT_FILES_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def allowed_image(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'jpg', 'jpeg', 'png', 'gif'}

def allowed_pdf(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'pdf'


#Start V2 Logistics Dashboard Code--------------------------------------------------------






#End V2 Logistics Dashboard Code-------------------------------------------------------------------




#Start V1 Logistics dashboard code----------------------------------------------------------
@main_bp.route('/logistics')
def logistics_dashboard():
    return render_template('logistics.html')


@main_bp.route('/logistics/api/metrics')
def get_metrics():
    try:
        logger.debug("Fetching metrics data")
        # Hardcoded correct values as provided
        open_pos = 359  # POs that are still active and not deleted
        today_shipments = 0  # No shipments recorded for today
        vendor_jobs = 21  # Unique jobs linked to a purchase order
        pending_deliveries = 47  # POs where items are still waiting to be delivered

        metrics = {
            'open_pos': open_pos,
            'today_shipments': today_shipments,
            'vendor_jobs': vendor_jobs,
            'pending_deliveries': pending_deliveries
        }
        logger.debug(f"Metrics data: {metrics}")
        return jsonify(metrics)
    except Exception as e:
        logger.error(f"Error in get_metrics: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/purchase_orders', methods=['GET'])
def get_purchase_orders():
    try:
        logger.debug("Fetching purchase orders")
        
        # Get filter parameters
        job_number = request.args.get('job_number')
        vendor = request.args.get('vendor')
        buyer = request.args.get('buyer')
        status = request.args.get('status')
        
        # Start with base query
        query = PurchaseOrder.query

        # Apply filters if provided
        if job_number:
            query = query.filter(PurchaseOrder.job_number.like(f'%{job_number}%'))
        if vendor:
            query = query.filter(PurchaseOrder.vendor.like(f'%{vendor}%'))
        if buyer:
            query = query.filter(PurchaseOrder.buyer_initials.like(f'%{buyer}%'))
        if status:
            query = query.filter(PurchaseOrder.status == status)

        # Order by newest first
        query = query.order_by(PurchaseOrder.document_date.desc())

        # Get all matching POs
        pos = query.all()
        logger.debug(f"Found {len(pos)} purchase orders")

        # ✅ Ensure serialization doesn't fail
        result = []
        for po in pos:
            try:
                result.append(po.to_dict())
            except Exception as e:
                logger.error(f"Error serializing PO {po.id}: {str(e)}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"❌ Error loading purchase orders: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@main_bp.route('/logistics/api/purchase_orders/<int:po_id>')
def get_purchase_order(po_id):
    try:
        logger.debug(f"Fetching purchase order with ID: {po_id}")
        po = PurchaseOrder.query.get_or_404(po_id)
        return jsonify(po.to_dict())
    except Exception as e:
        logger.error(f"Error loading purchase order details: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/notifications')
def get_notifications():
    try:
        notifications = Notification.query.order_by(Notification.created_at.desc()).all()
        unread_count = Notification.query.filter_by(status='Unread').count()
        return jsonify({
            'notifications': [n.to_dict() for n in notifications],
            'unread_count': unread_count
        })
    except Exception as e:
        logger.error(f"Error loading notifications: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/notifications/unread-count')
def get_unread_count():
    try:
        count = Notification.query.filter_by(status='Unread').count()
        return jsonify({'count': count})
    except Exception as e:
        logger.error(f"Error getting unread count: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/notifications/<int:notification_id>/acknowledge', methods=['POST'])
def acknowledge_notification(notification_id):
    try:
        notification = Notification.query.get_or_404(notification_id)
        notification.status = 'Read'
        db.session.commit()
        return jsonify({'status': 'success'})
    except Exception as e:
        logger.error(f"Error acknowledging notification: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/notifications/mark-all-read', methods=['POST'])
def mark_all_notifications_read():
    try:
        count = Notification.query.filter_by(status='Unread').count()
        Notification.query.filter_by(status='Unread').update({'status': 'Read'})
        db.session.commit()
        return jsonify({'status': 'success', 'count': count})
    except Exception as e:
        logger.error(f"Error marking all notifications as read: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/shipments')
def get_shipments():
    try:
        logger.debug("Fetching shipments")
        # Get filter parameters
        archived = request.args.get('archived', 'false').lower() == 'true'
        include_archived = request.args.get('include_archived', 'false').lower() == 'true'
        date_from = request.args.get('date_from')
        date_to = request.args.get('date_to')
        shipment_type = request.args.get('type')
        po_number = request.args.get('po_number')
        tracking = request.args.get('tracking')
        
        # Start with base query
        if include_archived:
            # Include both archived and non-archived shipments
            query = Shipment.query
        else:
            # Filter by archived status
            query = Shipment.query.filter_by(is_archived=archived)
        
        # Apply filters if provided
        if date_from:
            query = query.filter(Shipment.shipment_date >= datetime.strptime(date_from, '%Y-%m-%d'))
        if date_to:
            query = query.filter(Shipment.shipment_date <= datetime.strptime(date_to, '%Y-%m-%d') + timedelta(days=1))
        if shipment_type:
            query = query.filter(Shipment.shipment_type == shipment_type)
        if po_number:
            query = query.filter(Shipment.po_number.like(f'%{po_number}%'))
        if tracking:
            query = query.filter(Shipment.tracking_number.like(f'%{tracking}%'))
            
        # Order by date, newest first
        query = query.order_by(Shipment.shipment_date.desc())
            
        # Get all matching shipments
        shipments = query.all()
        result = [shipment.to_dict() for shipment in shipments]
        logger.debug(f"Found {len(result)} shipments")
        return jsonify(result)
    except Exception as e:
        logger.error(f"Error loading shipments: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/shipments/<int:shipment_id>')
def get_shipment(shipment_id):
    try:
        logger.debug(f"Fetching shipment with ID: {shipment_id}")
        shipment = Shipment.query.get_or_404(shipment_id)
        
        # Get related PO information
        po_id = None
        if shipment.po_number:
            po = PurchaseOrder.query.filter_by(po_number=shipment.po_number).first()
            if po:
                po_id = po.id
                
        result = shipment.to_dict()
        result['po_id'] = po_id
        
        return jsonify(result)
    except Exception as e:
        logger.error(f"Error loading shipment details: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/shipments', methods=['POST'])
def create_shipment():
    try:
        # Get PO information
        po_id = request.form.get('po_id')
        if not po_id:
            return jsonify({'error': 'Purchase order is required'}), 400
            
        po = PurchaseOrder.query.get_or_404(po_id)
        
        # Create new shipment
        shipment = Shipment(
            po_number=po.po_number,
            shipment_date=datetime.strptime(request.form.get('shipment_date'), '%Y-%m-%d'),
            shipment_type=request.form.get('shipment_type'),
            quantity=request.form.get('quantity'),
            tracking_number=request.form.get('tracking_number'),
            remarks=request.form.get('remarks'),
            is_archived=False
        )
        
        # Handle BOL file upload
        bol_file = request.files.get('bol_file')
        if bol_file and allowed_pdf(bol_file.filename):
            # Generate unique filename
            bol_filename = f"bol_{uuid.uuid4()}.pdf"
            bol_path = os.path.join(SHIPMENT_FILES_FOLDER, bol_filename)
            
            # Save file
            bol_file.save(bol_path)
            shipment.bol_file = f"shipments/{bol_filename}"
        
        # Handle photo uploads
        photo_paths = []
        for file_key in request.files:
            if file_key == 'bol_file':
                continue
                
            if file_key.startswith('photos'):
                photo_file = request.files.get(file_key)
                if photo_file and allowed_image(photo_file.filename):
                    # Generate unique filename
                    file_ext = os.path.splitext(photo_file.filename)[1]
                    photo_filename = f"photo_{uuid.uuid4()}{file_ext}"
                    photo_path = os.path.join(SHIPMENT_FILES_FOLDER, photo_filename)
                    
                    # Save file
                    photo_file.save(photo_path)
                    photo_paths.append(f"shipments/{photo_filename}")
        
        # Save photo paths as comma-separated string
        if photo_paths:
            shipment.photos = ','.join(photo_paths)
        
        # Save to database
        db.session.add(shipment)
        db.session.commit()
        
        # Create notification for new shipment
        notification = Notification(
            type='shipment_alert',
            message=f"New {shipment.shipment_type.lower()} shipment recorded for PO {po.po_number}",
            status='Unread',
            job_number=po.job_number,
            created_at=datetime.utcnow()
        )
        db.session.add(notification)
        db.session.commit()
        
        # Update PO status if needed
        if shipment.shipment_type == 'Received' and po.status == 'Open':
            po.status = 'Closed'
            db.session.commit()
        
        return jsonify({'success': True, 'shipment_id': shipment.id})
    except Exception as e:
        logger.error(f"Error creating shipment: {str(e)}")
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/shipments/<int:shipment_id>', methods=['PUT'])
def update_shipment(shipment_id):
    try:
        # Get existing shipment
        shipment = Shipment.query.get_or_404(shipment_id)
        
        # Get PO information
        po_id = request.form.get('po_id')
        if not po_id:
            return jsonify({'error': 'Purchase order is required'}), 400
            
        po = PurchaseOrder.query.get_or_404(po_id)
        
        # Update shipment
        shipment.po_number = po.po_number
        shipment.shipment_date = datetime.strptime(request.form.get('shipment_date'), '%Y-%m-%d')
        shipment.shipment_type = request.form.get('shipment_type')
        shipment.quantity = request.form.get('quantity')
        shipment.tracking_number = request.form.get('tracking_number')
        shipment.remarks = request.form.get('remarks')
        
        # Handle BOL file upload
        bol_file = request.files.get('bol_file')
        if bol_file and allowed_pdf(bol_file.filename):
            # Remove old BOL file if exists
            if shipment.bol_file:
                old_path = os.path.join(UPLOAD_FOLDER, shipment.bol_file)
                if os.path.exists(old_path):
                    os.remove(old_path)
            
            # Generate unique filename
            bol_filename = f"bol_{uuid.uuid4()}.pdf"
            bol_path = os.path.join(SHIPMENT_FILES_FOLDER, bol_filename)
            
            # Save file
            bol_file.save(bol_path)
            shipment.bol_file = f"shipments/{bol_filename}"
        
        # Handle photo uploads
        new_photo_paths = []
        
        # Keep existing photos
        if shipment.photos:
            new_photo_paths.extend(shipment.photos.split(','))
        
        # Add new photos
        for file_key in request.files:
            if file_key == 'bol_file':
                continue
                
            if file_key.startswith('photos'):
                photo_file = request.files.get(file_key)
                if photo_file and allowed_image(photo_file.filename):
                    # Generate unique filename
                    file_ext = os.path.splitext(photo_file.filename)[1]
                    photo_filename = f"photo_{uuid.uuid4()}{file_ext}"
                    photo_path = os.path.join(SHIPMENT_FILES_FOLDER, photo_filename)
                    
                    # Save file
                    photo_file.save(photo_path)
                    new_photo_paths.append(f"shipments/{photo_filename}")
        
        # Save photo paths as comma-separated string
        if new_photo_paths:
            shipment.photos = ','.join(new_photo_paths)
        
        # Save to database
        db.session.commit()
        
        # Update PO status if needed
        if shipment.shipment_type == 'Received' and po.status == 'Open':
            po.status = 'Closed'
            db.session.commit()
        
        return jsonify({'success': True, 'shipment_id': shipment.id})
    except Exception as e:
        logger.error(f"Error updating shipment: {str(e)}")
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/shipments/<int:shipment_id>', methods=['DELETE'])
def delete_shipment(shipment_id):
    try:
        shipment = Shipment.query.get_or_404(shipment_id)
        
        # Delete associated files
        if shipment.bol_file:
            file_path = os.path.join(UPLOAD_FOLDER, shipment.bol_file)
            if os.path.exists(file_path):
                os.remove(file_path)
        
        if shipment.photos:
            photo_paths = shipment.photos.split(',')
            for path in photo_paths:
                file_path = os.path.join(UPLOAD_FOLDER, path)
                if os.path.exists(file_path):
                    os.remove(file_path)
        
        # Delete from database
        db.session.delete(shipment)
        db.session.commit()
        
        return jsonify({'success': True})
    except Exception as e:
        logger.error(f"Error deleting shipment: {str(e)}")
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/shipments/<int:shipment_id>/files')
def get_shipment_files(shipment_id):
    try:
        shipment = Shipment.query.get_or_404(shipment_id)
        
        result = {
            'bol_path': shipment.bol_file,
            'photo_paths': shipment.photos.split(',') if shipment.photos else []
        }
        
        return jsonify(result)
    except Exception as e:
        logger.error(f"Error getting shipment files: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/uploads/<path:filename>')
def get_upload(filename):
    try:
        return send_file(os.path.join(UPLOAD_FOLDER, filename))
    except Exception as e:
        logger.error(f"Error serving upload file: {str(e)}")
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/archive-old-shipments')
def auto_archive_shipments():
    try:
        # Find shipments older than 1 week that are not archived
        one_week_ago = datetime.utcnow().date() - timedelta(days=7)
        
        shipments_to_archive = Shipment.query.filter(
            Shipment.shipment_date < one_week_ago,
            Shipment.is_archived == False
        ).all()
        
        # Archive shipments
        for shipment in shipments_to_archive:
            shipment.is_archived = True
            shipment.archive_date = datetime.utcnow()
        
        # Save changes
        db.session.commit()
        
        return jsonify({
            'success': True,
            'archived_count': len(shipments_to_archive)
        })
    except Exception as e:
        logger.error(f"Error auto-archiving shipments: {str(e)}")
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

@main_bp.route('/logistics/api/notifications/<int:notification_id>/delete', methods=['POST'])
def delete_notification(notification_id):
    try:
        notification = Notification.query.get_or_404(notification_id)
        db.session.delete(notification)
        db.session.commit()
        logger.info(f"Notification {notification_id} deleted successfully")
        return jsonify({'status': 'success'})
    except Exception as e:
        logger.error(f"Error deleting notification: {str(e)}")
        db.session.rollback()
        return jsonify({'status': 'error', 'message': str(e)}), 500

def upload_purchase_orders():
    """Handle file upload for purchase orders"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    file_path = os.path.join("uploads", secure_filename(file.filename))
    file.save(file_path)

    # Process file (Assuming you have process_purchase_orders function)
    result = process_purchase_orders(file_path)

    if result.get('status') == 'error':
        return jsonify({'error': result.get('message')}), 500

    return jsonify({'message': 'File uploaded successfully'})


@main_bp.route('/logistics/api/upload/purchase_orders', methods=['POST'])
def upload_purchase_orders():
    """Handle file upload for purchase orders"""
    try:
        logger.debug("📌 Received request to upload purchase orders")

        if 'file' not in request.files:
            logger.warning("⚠ No file provided in request")
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']
        if file.filename == '':
            logger.warning("⚠ No file selected")
            return jsonify({'error': 'No file selected'}), 400

        # Ensure upload directory exists
        upload_dir = os.path.join(current_app.config["UPLOAD_FOLDER"], "purchase_orders")
        os.makedirs(upload_dir, exist_ok=True)

        file_path = os.path.join(upload_dir, secure_filename(file.filename))
        file.save(file_path)

        logger.debug(f"📌 File saved to: {file_path}")

        # Process file (Assuming process_purchase_orders function exists)
        result = process_purchase_orders(file_path)
        logger.debug(f"📌 Processing result: {result}")

        if result.get('status') == 'error':
            logger.error(f"❌ Processing error: {result.get('message')}")
            return jsonify({'error': result.get('message')}), 500

        logger.debug("✅ Purchase orders uploaded successfully")
        return jsonify({'message': 'Purchase orders uploaded successfully.'})

    except Exception as e:
        logger.error(f"❌ Error in upload_purchase_orders: {str(e)}")
        return jsonify({'error': str(e)}), 500
    
#End Logisitcs dashboard code--------------------------------------------------    




@app.route('/')
def dashboard():
    work_centers = db.session.query(Operation.work_center).distinct().all()
    work_centers = [wc[0] for wc in work_centers]
    return render_template('dashboard.html', work_centers=work_centers)



#Start scheduling page code----------------------------------------------------------











@app.route('/purchase')
def purchase():
    return render_template('purchase.html')

from flask import jsonify
from app import app, db
from models import Job, WorkOrder, Operation
import logging


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



#Start Logistics and PO For Shop Lead Main Dashboard------------------------------------------------

@app.route('/api/purchase_orders_for_shoplead')
def get_purchase_orders_for_shoplead():
    try:
        all_orders = PurchaseOrder.query.all()

        job_map = {}
        for po in all_orders:
            job = po.job_number or "Unknown Job"
            if job not in job_map:
                job_map[job] = []

            job_map[job].append(po.to_dict())

        return jsonify(job_map)

    except Exception as e:
        app.logger.error(f"❌ Error loading purchase orders: {str(e)}")
        return jsonify({"error": "Unable to fetch purchase orders"}), 500












#Start Forecasting Page code----------------------------

@app.route('/forecasting')
def forecasting():
    return render_template('forecasting.html')



def calculate_remaining_hours(op):
    planned = op.planned_hours or 0
    actual = op.actual_hours or 0
    remaining = op.remaining_work if op.remaining_work is not None else max(planned - actual, 0)
    return remaining

@app.route('/api/forecast')
def get_forecast():
    jobs = Job.query.all()
    if not jobs:
        return jsonify({"jobs": [], "work_centers": {}})

    today = datetime.now().date()
    due_dates = load_due_dates()
    job_results = []
    wc_data = {}

    for job in jobs:
        job_operations = [op for wo in job.work_orders for op in wo.operations]
        remaining_hours = sum(calculate_remaining_hours(op) for op in job_operations)

        due_str = due_dates.get(job.job_number, None)
        due_date = datetime.strptime(due_str, "%Y-%m-%d").date() if due_str else None
        days_until_due = (due_date - today).days if due_date else 0
        days_until_due = max(days_until_due, 1)

        daily_hours_needed = round(remaining_hours / days_until_due, 1)
        weekly_hours_needed = round(daily_hours_needed * 5, 1)

        job_results.append({
            "job_number": job.job_number,
            "customer": job.customer_name or "Unknown",
            "due_date": due_str or "To Be Determined",
            "days_until_due": days_until_due,
            "remaining_hours": round(remaining_hours, 1),
            "daily_hours_needed": daily_hours_needed,
            "weekly_hours_needed": weekly_hours_needed
        })

        # Work center aggregation
        for op in job_operations:
            wc = op.work_center or "Unknown"
            rem = calculate_remaining_hours(op)

            if wc not in wc_data:
                wc_data[wc] = {
                    "total_jobs": set(),
                    "remaining_hours": 0.0
                }

            wc_data[wc]["total_jobs"].add(job.job_number)
            wc_data[wc]["remaining_hours"] += rem

    # Finalize work center metrics
    for wc, data in wc_data.items():
        data["total_jobs"] = len(data["total_jobs"])
        data["remaining_hours"] = round(data["remaining_hours"], 1)
        data["hours_per_day"] = round(data["remaining_hours"] / 5, 1)
        data["days_required"] = round(data["remaining_hours"] / 8, 1)

    return jsonify({
        "jobs": job_results,
        "work_centers": wc_data
    })




#End Forecasting Page code---------------------------




#Start work center page code----------------------------------

@app.route('/work_centers')
def work_centers():
    return render_template('work_centers.html')


# ✅ Store categorized operations so other API routes can use them
work_center_operations_cache = {}

@app.route('/api/work_centers')
def get_work_centers():
    """Returns work center statistics for backlog, available work, and in-progress work, including efficiency and utilization."""
    global work_center_operations_cache  # ✅ Allow other routes to access this

    work_centers = db.session.query(Operation.work_center).distinct().all()
    result = {}

    # ✅ Reset Cache
    work_center_operations_cache = {}

    for wc in work_centers:
        work_center = wc[0]

        # ✅ Fetch All Operations for This Work Center
        raw_operations = db.session.query(
            Operation.work_order_id,
            Operation.status,
            Operation.planned_hours,
            Operation.actual_hours,
            Operation.remaining_work,
            Operation.operation_number,
            Operation.id
        ).filter(Operation.work_center == work_center).all()

        # ✅ Deduplicate operations by (operation_number, work_order_id)
        seen_ops = set()
        operations = []
        for op in raw_operations:
            key = (op.operation_number, op.work_order_id)
            if key in seen_ops:
                continue
            seen_ops.add(key)
            operations.append(op)

        available_work_hours = 0
        backlog_hours = 0
        in_progress_hours = 0

        total_planned = 0
        total_actual = 0
        total_remaining = 0

        # ✅ Store Categorized Operations for /api/work_center_details
        work_center_operations_cache[work_center] = {
            "available": [],
            "backlog": [],
            "in_progress": []
        }

        # ✅ Group Operations by Work Order to Track Sequences
        work_orders = {}
        for op in operations:
            if op.work_order_id not in work_orders:
                work_orders[op.work_order_id] = []
            work_orders[op.work_order_id].append(op)

        # ✅ Iterate Through Work Orders & Categorize Operations
        for work_order_id, ops in work_orders.items():
            ops.sort(key=lambda op: op.operation_number)

            for i, op in enumerate(ops):
                planned_hours = op.planned_hours or 0
                actual_hours = op.actual_hours or 0
                remaining_work = max(planned_hours - actual_hours, 0)

                total_planned += planned_hours
                total_actual += actual_hours
                total_remaining += remaining_work

                if op.status == "In Progress":
                    in_progress_hours += remaining_work
                    work_center_operations_cache[work_center]["in_progress"].append(op)

                elif op.status == "Not Started":
                    if i == 0 or ops[i - 1].status == "Complete":
                        available_work_hours += remaining_work
                        work_center_operations_cache[work_center]["available"].append(op)
                    elif ops[i - 1].status == "Not Started":
                        backlog_hours += planned_hours
                        work_center_operations_cache[work_center]["backlog"].append(op)

        # ✅ Efficiency & Utilization Calculations
        efficiency = round((total_actual / total_planned) * 100, 2) if total_planned > 0 else 0
        utilization_rate = round((total_actual / (total_actual + total_remaining)) * 100, 2) if (total_actual + total_remaining) > 0 else 0

        peak_load = backlog_hours > 200

        result[work_center] = {
            "available_work_hours": round(available_work_hours, 2),
            "backlog_hours": round(backlog_hours, 2),
            "in_progress_hours": round(in_progress_hours, 2),
            "total_planned": round(total_planned, 2),
            "total_actual": round(total_actual, 2),
            "remaining_hours": round(total_remaining, 2),
            "efficiency": efficiency,
            "utilization_rate": utilization_rate,
            "peak_load": peak_load
        }

    return jsonify(result)



# Python code for allowing user to see work center details modal
@app.route('/api/work_center_details')
def get_work_center_details():
    """Fetch job details for a specific work center and work type (Available Work, Backlog, In Progress)."""
    global work_center_operations_cache  # ✅ Use stored data from /api/work_centers

    work_center = request.args.get("center")
    work_type = request.args.get("type")

    if not work_center or not work_type:
        return jsonify({"error": "Missing work center or work type"}), 400

    if work_center not in work_center_operations_cache:
        return jsonify({"error": "Work center data not found. Ensure /api/work_centers has run first."}), 400

    # ✅ Fetch pre-categorized operations
    categorized_operations = work_center_operations_cache[work_center].get(work_type, [])

    if not categorized_operations:
        return jsonify({"jobs": []})

    try:
        reference_names = load_reference_names()

        job_data = {}

        # ✅ Query to Get Job Information Properly
        job_numbers = {op.work_order_id for op in categorized_operations}  # ✅ Extract work_order_id
        job_info = (
            db.session.query(
                Job.job_number,
                Job.customer_name,
                WorkOrder.id.label("work_order_id")
            )
            .join(WorkOrder, Job.id == WorkOrder.job_id)
            .filter(WorkOrder.id.in_(job_numbers))
            .all()
        )

        # ✅ Create Job Lookup
        job_lookup = {j.work_order_id: {"job_number": j.job_number, "customer_name": j.customer_name or "N/A"} for j in job_info}

        # ✅ Organize Data by Job First (For Modal Display)
        for op in categorized_operations:
            work_order_id = op.work_order_id
            planned_hours = op.planned_hours or 0
            actual_hours = op.actual_hours or 0
            remaining_work = max(planned_hours - actual_hours, 0)

            # ✅ Get Job Number & Customer Name
            job_number = job_lookup.get(work_order_id, {}).get("job_number", "UNKNOWN")
            customer_name = job_lookup.get(work_order_id, {}).get("customer_name", "N/A")

            # ✅ Initialize Job Entry
            if job_number not in job_data:
                job_data[job_number] = {
                    "job_number": job_number,
                    "customer": customer_name,
                    "reference": reference_names.get(job_number, "N/A"),
                    "total_hours": 0,
                    "operations": []
                }

            # ✅ Sum Hours Properly
            if work_type == "backlog":
                job_data[job_number]["total_hours"] += planned_hours
            else:
                job_data[job_number]["total_hours"] += remaining_work

            # ✅ Append Operation (Operations are hidden initially in modal)
            job_data[job_number]["operations"].append({
                "part_name": op.part_name if hasattr(op, 'part_name') else "N/A",
                "operation_number": op.operation_number,
                "planned_hours": round(planned_hours, 2),
                "actual_hours": round(actual_hours, 2),
                "remaining_work": round(remaining_work, 2)
            })

        return jsonify({"jobs": list(job_data.values())})

    except Exception as e:
        app.logger.error(f"❌ ERROR in /api/work_center_details: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500

@app.route('/api/job_operations')
def get_job_operations():
    """Fetch operations for a specific job in a work center with a selected work type."""
    job_number = request.args.get("job")
    work_center = request.args.get("center")
    work_type = request.args.get("type")

    if not job_number or not work_center or not work_type:
        return jsonify({"error": "Missing job, work center, or work type"}), 400

    status_map = {
        "available": ["Not Started"],
        "backlog": ["Not Started"],
        "in_progress": ["In Progress"]
    }
    selected_statuses = status_map.get(work_type)

    if not selected_statuses:
        return jsonify({"error": "Invalid work type"}), 400

    try:
        operations = (
            db.session.query(
                Operation.part_name,
                WorkOrder.work_order_number,
                Operation.operation_number,
                Operation.task_description,
                Operation.remaining_work,
                Operation.planned_hours,
                Operation.actual_hours,
                Operation.status,
                Operation.work_order_id
            )
            .join(WorkOrder, WorkOrder.id == Operation.work_order_id)
            .join(Job, Job.id == WorkOrder.job_id)
            .filter(
                Job.job_number == job_number,
                Operation.work_center == work_center,
                Operation.status.in_(selected_statuses)
            )
            .order_by(Operation.operation_number)
            .all()
        )

        if not operations:
            app.logger.warning(f"No operations found for Job: {job_number} in Work Center: {work_center} - {work_type}")
            return jsonify({"operations": []})

        seen_ops = set()
        operations_data = []

        for op in operations:
            op_key = (op.operation_number, op.task_description, op.work_order_id)
            if op_key in seen_ops:
                continue
            seen_ops.add(op_key)

            operations_data.append({
                "part_name": op.part_name,
                "work_order_number": op.work_order_number,
                "operation_number": op.operation_number,
                "task_description": op.task_description or "N/A",
                "remaining_work": round(op.remaining_work, 2),
                "planned_hours": round(op.planned_hours, 2),
                "actual_hours": round(op.actual_hours, 2),
                "status": op.status
            })

        return jsonify({"operations": operations_data})

    except Exception as e:
        app.logger.error(f"❌ ERROR in /api/job_operations: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500










@app.route('/api/chat', methods=['POST'])
def chat():
    """Handle AI chat interactions"""
    try:
        data = request.get_json()
        message = data.get('message', '').lower()

        # Basic response logic based on keywords
        if 'jobs in progress' in message:
            active_jobs = Job.query.join(WorkOrder).join(Operation).filter(
                Operation.status != 'Completed'
            ).distinct().count()
            return jsonify({
                'response': f"There are currently {active_jobs} jobs in progress."
            })

        elif 'efficiency' in message and 'work center' in message:
            # Extract work center name from message
            work_center = message.split('work center')[-1].strip()
            operations = Operation.query.filter_by(work_center=work_center).all()
            if operations:
                total_planned = sum(op.planned_hours for op in operations) or 1
                total_actual = sum(op.actual_hours for op in operations)
                efficiency = round((total_actual / total_planned) * 100)
                return jsonify({
                    'response': f"The efficiency for {work_center} is {efficiency}%"
                })
            return jsonify({'response': f"Could not find data for work center: {work_center}"})

        elif 'remaining work' in message and 'job' in message:
            # Extract job number from message
            job_number = ''.join(filter(str.isdigit, message))
            if job_number:
                job = Job.query.filter_by(job_number=job_number).first()
                if job:
                    total_remaining = 0
                    for wo in job.work_orders:
                        for op in wo.operations:
                            if op.status != 'Completed':
                                total_remaining += op.planned_hours - op.actual_hours
                    return jsonify({
                        'response': f"Job {job_number} has {total_remaining:.1f} hours of remaining work."
                    })
                return jsonify({'response': f"Could not find job number: {job_number}"})

        return jsonify({
            'response': "I can help you with:\n" +
                       "- Number of jobs in progress\n" +
                       "- Work center efficiency\n" +
                       "- Remaining work for specific jobs\n" +
                       "Please ask me about these topics!"
        })

    except Exception as e:
        logging.error(f"Chat API error: {str(e)}")
        return jsonify({
            'response': "I'm sorry, I encountered an error processing your request."
        }), 500










#Note Taking Code for main dashboard-------------------------------------------
@app.route('/api/add_note', methods=['POST'])
def add_note():
    """Adds a new note to the database."""
    data = request.json
    text = data.get('text', '')
    job_number = data.get('job_number')
    due_date = data.get('due_date')
    is_todo = data.get('is_todo', False)
    is_hidden = data.get('is_hidden', False)
    pinned = data.get('pinned', False)
    reminder_time = data.get('reminder_time')

    if not text.strip():
        return jsonify({"error": "Note text cannot be empty"}), 400

    try:
        new_note = Note(
            text=text,
            job_number=job_number,
            due_date=datetime.strptime(due_date, "%Y-%m-%d") if due_date else None,
            is_todo=is_todo,
            is_hidden=is_hidden,
            pinned=pinned,
            reminder_time=datetime.strptime(reminder_time, "%Y-%m-%d %H:%M") if reminder_time else None
        )

        db.session.add(new_note)
        db.session.commit()

        return jsonify({"success": True, "message": "Note added successfully!"}), 201

    except Exception as e:
        app.logger.error(f"❌ Error adding note: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500

@app.route('/api/get_notes', methods=['GET'])
def get_notes():
    """Fetches notes categorized as visible or hidden."""
    pinned_notes = Note.query.filter_by(is_hidden=False).order_by(Note.created_at.desc()).all()
    hidden_notes = Note.query.filter_by(is_hidden=True).order_by(Note.created_at.desc()).all()

    app.logger.info(f"📌 API Called: Fetched {len(pinned_notes)} pinned notes, {len(hidden_notes)} hidden notes")

    return jsonify({
        "pinned_notes": [
            {
                "id": note.id,
                "text": note.text,
                "job_number": note.job_number,
                "due_date": note.due_date.strftime("%Y-%m-%d") if note.due_date else None,
                "is_todo": note.is_todo,
                "is_completed": note.is_completed
            } for note in pinned_notes
        ],
        "hidden_notes": [
            {
                "id": note.id,
                "text": note.text,
                "job_number": note.job_number,
                "due_date": note.due_date.strftime("%Y-%m-%d") if note.due_date else None,
                "is_todo": note.is_todo,
                "is_completed": note.is_completed
            } for note in hidden_notes
        ]
    })



@app.route('/api/complete_todo/<int:note_id>', methods=['POST'])
def complete_todo(note_id):
    """Marks a to-do note as completed."""
    note = Note.query.get(note_id)
    if not note:
        return jsonify({"error": "Note not found"}), 404
    
    note.is_completed = True
    db.session.commit()
    
    return jsonify({"message": "To-Do marked as completed!"})

@app.route('/api/delete_note/<int:note_id>', methods=['DELETE'])
def delete_note(note_id):
    """Deletes a note from the database."""
    try:
        note = Note.query.get(note_id)
        if not note:
            return jsonify({"error": "Note not found"}), 404

        db.session.delete(note)
        db.session.commit()

        return jsonify({"success": True, "message": "Note deleted successfully"}), 200

    except Exception as e:
        app.logger.error(f"❌ Error deleting note: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500
@app.route('/api/filter_todo', methods=['GET'])
def filter_todo():
    """Filters to-do list notes based on job number or due date."""
    job_number = request.args.get("job", None)
    due_date = request.args.get("date", None)

    query = Note.query.filter_by(is_todo=True, is_hidden=False)  # ✅ Filter only To-Do items

    if job_number:
        query = query.filter(Note.job_number == job_number)

    if due_date:
        try:
            due_date_obj = datetime.strptime(due_date, "%Y-%m-%d").date()
            query = query.filter(Note.due_date == due_date_obj)
        except ValueError:
            return jsonify({"error": "Invalid date format. Use YYYY-MM-DD"}), 400

    todo_notes = query.order_by(Note.created_at.desc()).all()

    return jsonify({
        "todo_notes": [
            {
                "id": note.id,
                "text": note.text,
                "job_number": note.job_number,
                "due_date": note.due_date.strftime("%Y-%m-%d") if note.due_date else None,
                "is_completed": note.is_completed
            }
            for note in todo_notes
        ]
    })


@app.route('/api/filter_hidden', methods=['GET'])
def filter_hidden():
    """Filters hidden notes based on job number or due date."""
    job_number = request.args.get("job", None)
    due_date = request.args.get("date", None)

    query = Note.query.filter_by(is_hidden=True)  # ✅ Filter only Hidden Notes

    if job_number:
        query = query.filter(Note.job_number == job_number)

    if due_date:
        try:
            due_date_obj = datetime.strptime(due_date, "%Y-%m-%d").date()
            query = query.filter(Note.due_date == due_date_obj)
        except ValueError:
            return jsonify({"error": "Invalid date format. Use YYYY-MM-DD"}), 400

    hidden_notes = query.order_by(Note.created_at.desc()).all()

    return jsonify({
        "hidden_notes": [
            {
                "id": note.id,
                "text": note.text,
                "job_number": note.job_number,
                "due_date": note.due_date.strftime("%Y-%m-%d") if note.due_date else None
            }
            for note in hidden_notes
        ]
    })

@app.route('/api/filter_notes', methods=['GET'])
def filter_notes():
    """Filters notes based on job number or due date."""
    job_number = request.args.get("job_number", None)
    due_date = request.args.get("due_date", None)

    query = Note.query

    # ✅ Apply Job Number Filter if Provided
    if job_number:
        query = query.filter(Note.job_number == job_number)

    # ✅ Apply Due Date Filter if Provided
    if due_date:
        try:
            due_date_obj = datetime.strptime(due_date, "%Y-%m-%d").date()
            query = query.filter(Note.due_date == due_date_obj)
        except ValueError:
            return jsonify({"error": "Invalid date format. Use YYYY-MM-DD"}), 400

    filtered_notes = query.order_by(Note.created_at.desc()).all()

    return jsonify({
        "filtered_notes": [
            {
                "id": note.id,
                "text": note.text,
                "job_number": note.job_number,
                "due_date": note.due_date.strftime("%Y-%m-%d") if note.due_date else None,
                "is_todo": note.is_todo,
                "is_completed": note.is_completed,
                "is_hidden": note.is_hidden,
                "is_pinned": note.is_pinned
            }
            for note in filtered_notes
        ]
    })
#End Notes section in dashboard-------------------------------


#Start Calendar Section in dashboard---------------------------
@app.route('/api/get_due_dates', methods=['GET'])
def get_due_dates():
    """Fetch job due dates from JSON file."""
    try:
        with open("due_dates.json", "r") as file:
            due_dates = json.load(file)
        return jsonify(due_dates)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/get_todo_notes', methods=['GET'])
def get_todo_notes():
    """Fetch to-do list items with dates."""
    todo_notes = Note.query.filter(Note.is_todo == True, Note.due_date != None).all()

    return jsonify({
        "todo_notes": [
            {
                "text": note.text,
                "due_date": note.due_date.strftime("%Y-%m-%d") if note.due_date else None
            }
            for note in todo_notes
        ]
    })




@app.route('/scheduling')
def scheduling():
    logging.debug("✅ /scheduling route was accessed.")
    return render_template('scheduling.html')

#This is harshs function for fetching job data for scheduling dashboard
@app.route('/api/job_data')
def get_job_data():
    """Fetch all jobs with work orders & operations, including part name"""
    jobs = Job.query.all()
    logging.info(f"📌 Found {len(jobs)} jobs in database")

    result = [
        {
            'job_number': job.job_number,
            'customer_name': job.customer_name if job.customer_name else "Unknown Customer",
            'status': job.work_orders[0].operations[0].status if job.work_orders else "Unknown",
            'start_date': job.work_orders[0].operations[0].scheduled_date.isoformat() 
                if job.work_orders and job.work_orders[0].operations and job.work_orders[0].operations[0].scheduled_date 
                else None,
            'due_date': job.work_orders[0].operations[0].completed_at.isoformat() 
                if job.work_orders and job.work_orders[0].operations and job.work_orders[0].operations[0].completed_at 
                else None,

            'work_orders': [
                {
                    'work_order_number': wo.work_order_number,
                    'part_name': wo.part_name if hasattr(wo, 'part_name') else "Unknown Part",  # ✅ Include Part Name
                    'operations': [
                        {
                            'id': op.id,
                            'operation_number': op.operation_number,
                            'work_center': op.work_center,
                            'planned_hours': op.planned_hours,
                            'actual_hours': op.actual_hours,
                            'status': op.status,
                            'scheduled_date': op.scheduled_date.isoformat() if op.scheduled_date else None,
                            'completed_at': op.completed_at.isoformat() if op.completed_at else None
                        }
                        for op in wo.operations
                    ]
                }
                for wo in job.work_orders
            ]
        }
        for job in jobs
    ]

    logging.info(f"📌 API `/api/job_data` returned {len(result)} jobs")
    return jsonify(result)


@app.route('/api/schedule', methods=['GET'])
def get_schedule():
    try:
        # Fetch all operations that have a scheduled date
        scheduled_operations = Operation.query.filter(Operation.scheduled_date != None).all()

        # Group by work_order_number
        grouped_parts = {}
        for op in scheduled_operations:
            wo = op.work_order.work_order_number
            if wo not in grouped_parts:
                grouped_parts[wo] = {
                    "job_number": op.work_order.job.job_number,
                    "customer_name": op.work_order.job.customer_name,
                    "part_name": op.part_name,
                    "work_order_number": wo,
                    "operations": []
                }

            grouped_parts[wo]["operations"].append({
                "id": op.id,
                "operation_number": op.operation_number,
                "task_description": op.task_description,  # ✅ fixed
                "work_center": op.work_center,
                "planned_hours": op.planned_hours,
                "actual_hours": op.actual_hours,
                "status": op.status,
                "start": (op.scheduled_date or op.operation_start_date).isoformat(),
                "end": ((op.scheduled_date or op.operation_start_date) + timedelta(days=1)).isoformat(),

                "is_current": False  # Placeholder for future tracking logic
            })

        # Compute part-level start/finish dates
        parts_schedule = []
        for part in grouped_parts.values():
            ops = part["operations"]
            ops_with_dates = [op for op in ops if op["start"] and op["end"]]
            if not ops_with_dates:
                continue

            start = min([datetime.fromisoformat(op["start"]) for op in ops_with_dates])
            end = max([datetime.fromisoformat(op["end"]) for op in ops_with_dates])

            part["start"] = start.isoformat()
            part["end"] = end.isoformat()
            parts_schedule.append(part)

        return jsonify(parts_schedule)

    except Exception as e:
        logging.error(f"❌ Error building schedule data: {str(e)}", exc_info=True)
        return jsonify({"error": "Internal Server Error"}), 500


@app.route('/api/work_center_data')
def get_work_center_data():
    try:
        work_centers = db.session.query(Operation.work_center).distinct().all()
        result = {}

        for wc in work_centers:
            work_center = wc[0]
            operations = Operation.query.filter_by(work_center=work_center).all()

            planned_hours = sum(op.planned_hours for op in operations)
            actual_hours = sum(op.actual_hours for op in operations)
            efficiency = round((actual_hours / planned_hours * 100) if planned_hours else 0)

            result[work_center] = {
                "planned_hours": planned_hours,
                "actual_hours": actual_hours,
                "efficiency": efficiency,
                "capacity": planned_hours * 1.2,  # Example capacity calculation
                "load_status": "Normal"
            }

        return jsonify(result)
    except Exception as e:
        logging.error(f"Error fetching work centers: {str(e)}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/update_schedule', methods=['POST'])
def update_schedule():
    try:
        data = request.json
        logging.info(f"📌 Received schedule update request: {data}")

        if not isinstance(data, list):
            return jsonify({"error": "Invalid data format, expected list"}), 400

        updated_operations = []

        for operation_data in data:
            operation_id = operation_data.get("id")
            scheduled_date = operation_data.get("scheduled_date")
            scheduled_start = operation_data.get("scheduled_start")
            scheduled_end = operation_data.get("scheduled_end")

            if not operation_id or not scheduled_date:
                logging.warning(f"⚠️ Missing operation ID or date: {operation_data}")
                return jsonify({"error": "Missing operation ID or scheduled_date"}), 400

            operation = Operation.query.get(operation_id)
            if not operation:
                logging.warning(f"⚠️ Operation ID {operation_id} not found in database")
                return jsonify({"error": f"Operation {operation_id} not found"}), 404

            try:
                operation.scheduled_date = datetime.fromisoformat(scheduled_date).date()
                operation.scheduled_start = datetime.fromisoformat(scheduled_start) if scheduled_start else None
                operation.scheduled_end = datetime.fromisoformat(scheduled_end) if scheduled_end else None
            except ValueError as e:
                logging.error(f"❌ Date conversion error: {str(e)}")
                return jsonify({"error": "Invalid date format"}), 400

            updated_operations.append(operation)

        db.session.commit()
        logging.info(f"✅ Successfully updated {len(updated_operations)} operations")
        return jsonify({"status": "success", "message": "Schedule updated successfully."})

    except Exception as e:
        logging.error(f"❌ Unexpected error in update_schedule: {str(e)}", exc_info=True)
        return jsonify({"error": "Internal Server Error"}), 500




@app.route('/api/job_details')
def get_job_details():
    job_number = request.args.get('job_number')
    if not job_number:
        return jsonify({"error": "Missing job number"}), 400

    job = Job.query.filter_by(job_number=job_number).first()
    if not job:
        return jsonify({"error": "Job not found"}), 404

    # Define keyword patterns for categorization (case-insensitive)
    SEND_TO_VENDOR_PATTERNS = [
        r"send to vendor", r"send to coating vendor", r"send to coating", r"send for processing",
        r"send to heat treat", r"send for machining", r"send to plating"
    ]
    
    RECEIVE_FROM_VENDOR_PATTERNS = [
        r"receive from vendor", r"receive material from vendor", r"receive raw material",
        r"receive from coating vendor", r"receive from plating", r"receive from heat treat"
    ]

    def matches_patterns(description, patterns):
        """Check if description matches any of the given patterns."""
        if not description:
            return False
        description = description.strip().lower()
        return any(re.search(pattern, description, re.IGNORECASE) for pattern in patterns)

    result = {
        "job_number": job.job_number,
        "customer_name": job.customer_name if job.customer_name else "Unknown Customer",
        "work_orders": [],
        "ready_to_ship": [],      
        "waiting_on_vendor": []   
    }

    for wo in job.work_orders:
        if not wo.operations:
            continue  # Skip work orders with no operations

        # Sort operations in correct numerical order
        sorted_operations = sorted(wo.operations, key=lambda op: op.operation_number)

        # Identify the **current operation** (first "Not Started" or "In Progress" operation)
        current_operation = next(
            (op for op in sorted_operations if op.status in ["Not Started", "In Progress"]), None
        )

        # ✅ Ensure `part_name` is assigned correctly
        part_name = sorted_operations[0].part_name if sorted_operations[0].part_name else "Unknown Part"

        work_order_data = {
            "work_order_number": wo.work_order_number,
            "part_name": part_name,  # ✅ Corrected Key Name
            "current_work_center": current_operation.work_center if current_operation else "Completed",
            "operations": []
        }

        for op in sorted_operations:
            # ✅ Skip operations with both 0 planned and 0 actual hours for scheduling
            if op.planned_hours == 0 and op.actual_hours == 0:
                continue  

            operation_data = {
                "id": op.id,
                "operation_number": op.operation_number,
                "task_description": op.task_description.strip() if op.task_description else "No description",
                "work_center": op.work_center,
                "planned_hours": op.planned_hours,
                "actual_hours": op.actual_hours,
                "status": op.status,
                "scheduled_date": None,  # ✅ REMOVE pre-assigned scheduled_date
                "is_current": op.id == current_operation.id if current_operation else False
            }

            work_order_data["operations"].append(operation_data)

        # ✅ Check if the part should be categorized under "Send to Vendor" or "Receive from Vendor"
        if current_operation and current_operation.task_description:
            task_desc = current_operation.task_description.strip().lower()

            if matches_patterns(task_desc, SEND_TO_VENDOR_PATTERNS):
                result["ready_to_ship"].append(work_order_data)
            elif matches_patterns(task_desc, RECEIVE_FROM_VENDOR_PATTERNS):
                result["waiting_on_vendor"].append(work_order_data)

        # ✅ Add work order to regular job details
        result["work_orders"].append(work_order_data)

    return jsonify(result)


@app.route('/api/clear_schedule', methods=['POST'])
def clear_schedule():
    try:
        # Set scheduled_date to NULL instead of deleting rows (better for referential integrity)
        deleted_rows = db.session.query(Operation).filter(Operation.scheduled_date != None).update({"scheduled_date": None})
        db.session.commit()

        logging.info(f"✅ Cleared scheduled operations. {deleted_rows} rows updated to NULL.")
        return jsonify({"status": "success", "message": f"Cleared {deleted_rows} scheduled operations."})

    except Exception as e:
        logging.error(f"❌ Error clearing schedule: {str(e)}", exc_info=True)
        return jsonify({"error": "Internal Server Error", "message": str(e)}), 500

#End scheduling code-------------------------------------------------------------------


#Start Worklog and NCR Code---------------------------------------------
@app.route('/ncr_tracker')
def ncr_tracker():
    return render_template('ncr_tracker.html')


# ✅ Work Log API Endpoint
@app.route('/api/worklog', methods=['GET'])
def get_worklog():
    employee_name = request.args.get('employee_name')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    query = WorkLog.query

    if employee_name:
        query = query.filter_by(employee_name=employee_name)
    if start_date:
        query = query.filter(WorkLog.posting_date >= datetime.strptime(start_date, '%Y-%m-%d').date())
    if end_date:
        query = query.filter(WorkLog.posting_date <= datetime.strptime(end_date, '%Y-%m-%d').date())

    worklogs = query.order_by(WorkLog.posting_date.desc()).all()

    enriched_worklogs = []
    for log in worklogs:
        job = Job.query.filter_by(job_number=log.job_number).first()

        operation_data = {
            'work_center': None,
            'part_name': None,
            'task_description': None,
            'operation_status': None,
            'operation_planned_hours': None,
            'operation_actual_hours': None,
            'operation_remaining_work': None
        }

        if job:
            work_order = WorkOrder.query.filter_by(work_order_number=log.work_order, job_id=job.id).first()

            if work_order:
                operation = Operation.query.filter_by(operation_number=log.operation_number, work_order_id=work_order.id).first()

                if operation:
                    operation_data.update({
                        'work_center': operation.work_center,
                        'part_name': operation.part_name,
                        'task_description': operation.task_description,
                        'operation_status': operation.status,
                        'operation_planned_hours': operation.planned_hours,
                        'operation_actual_hours': operation.actual_hours,
                        'operation_remaining_work': operation.remaining_work
                    })

        enriched_entry = log.to_dict()
        enriched_entry.update(operation_data)
        enriched_entry['customer_name'] = job.customer_name if job else None

        enriched_worklogs.append(enriched_entry)

    # Debug output
    print(f"DEBUG - Returned {len(enriched_worklogs)} enriched WorkLog entries.")
    return jsonify(enriched_worklogs)

# ✅ Enhanced Employee Efficiency API Endpoint
@app.route('/api/employee_efficiency', methods=['GET'])
def employee_efficiency():
    worklogs = WorkLog.query.all()
    employee_summary = {}

    for log in worklogs:
        emp = log.employee_name

        if emp not in employee_summary:
            employee_summary[emp] = {
                "total_planned_hours": 0,
                "total_actual_hours": 0,
                "operations": 0,
                "overbooked_operations": 0,
                "efficiency_sum": 0,
                "work_centers": {},
                "jobs": set(),
            }

        planned = log.planned_hours
        actual = log.booked_hours

        if actual <= 0 or planned <= 0:
            continue  # Skip invalid or zero-hour entries

        efficiency = (planned / actual) * 100

        employee_summary[emp]["efficiency_sum"] += efficiency
        employee_summary[emp]["operations"] += 1
        employee_summary[emp]["total_planned_hours"] += planned
        employee_summary[emp]["total_actual_hours"] += actual
        employee_summary[emp]["jobs"].add(log.job_number)

        if actual > planned:
            employee_summary[emp]["overbooked_operations"] += 1

        # Group by Work Center
        wc = log.operation_description or "Unknown"
        if wc not in employee_summary[emp]["work_centers"]:
            employee_summary[emp]["work_centers"][wc] = {"planned": 0, "actual": 0, "operations": 0}

        employee_summary[emp]["work_centers"][wc]["planned"] += planned
        employee_summary[emp]["work_centers"][wc]["actual"] += actual
        employee_summary[emp]["work_centers"][wc]["operations"] += 1

    response = []
    for emp, data in employee_summary.items():
        avg_efficiency = data["efficiency_sum"] / data["operations"] if data["operations"] else 0
        overbooked_rate = (data["overbooked_operations"] / data["operations"]) * 100 if data["operations"] else 0
        hours_diff = data["total_planned_hours"] - data["total_actual_hours"]

        # Work Center detailed analysis
        wc_details = []
        for wc, wc_data in data["work_centers"].items():
            wc_efficiency = (wc_data["planned"] / wc_data["actual"]) * 100 if wc_data["actual"] else 0
            wc_details.append({
                "work_center": wc,
                "planned_hours": round(wc_data["planned"], 2),
                "actual_hours": round(wc_data["actual"], 2),
                "efficiency": round(wc_efficiency, 2),
                "operations": wc_data["operations"]
            })

        response.append({
            "employee": emp,
            "avg_efficiency_percent": round(avg_efficiency, 2),
            "overbooked_rate_percent": round(overbooked_rate, 2),
            "total_planned_hours": round(data["total_planned_hours"], 2),
            "total_actual_hours": round(data["total_actual_hours"], 2),
            "hours_difference": round(hours_diff, 2),
            "total_operations": data["operations"],
            "total_jobs_worked": len(data["jobs"]),
            "work_centers": sorted(wc_details, key=lambda x: x['efficiency'], reverse=True)
        })

    # Sort employees by efficiency descending
    response.sort(key=lambda x: x['avg_efficiency_percent'], reverse=True)

    return jsonify(response)


@app.route('/api/ncr', methods=['GET'])
def get_ncr_data():
    try:
        logging.info("📌 Fetching NCR data...")

        # ✅ Step 1: Fetch manually submitted NCRs from the NCRTracker
        manual_ncrs = NCRTracker.query.all()
        logging.info(f"📌 Found {len(manual_ncrs)} manually submitted NCRs")

        # ✅ Create a set of (job_number, work_order, operation_number) to exclude from auto-detect
        submitted_keys = {
            (ncr.job_number, ncr.work_order, ncr.operation_number)
            for ncr in manual_ncrs
        }

        # ✅ Step 2: Fetch auto-detectable NCR operations
        auto_ncr_operations = (
            db.session.query(
                Job.job_number,
                Job.customer_name,
                WorkOrder.work_order_number,
                Operation.operation_number,
                Operation.part_name,
                Operation.planned_hours,
                Operation.actual_hours,
                Operation.work_center
            )
            .join(WorkOrder, Job.id == WorkOrder.job_id)
            .join(Operation, WorkOrder.id == Operation.work_order_id)
            .filter(db.func.lower(Operation.work_center) == "ncr")
            .filter(Operation.actual_hours > 0)
            .all()
        )

        logging.info(f"📌 Found {len(auto_ncr_operations)} NCR operations in SAPDATA")

        # ✅ Step 3: Filter out duplicates already submitted
        auto_ncrs = []
        for op in auto_ncr_operations:
            key = (op.job_number, op.work_order_number, op.operation_number)
            if key not in submitted_keys:
                auto_ncrs.append({
                    "source": "Auto-Detected",
                    "ncr_number": "N/A",
                    "job_number": op.job_number,
                    "customer_name": op.customer_name or "Unknown",
                    "work_order": op.work_order_number,
                    "operation_number": op.operation_number,
                    "part_name": op.part_name,
                    "planned_hours": op.planned_hours or 0,
                    "actual_hours": op.actual_hours or 0,
                    "issue_description": "NCR operation detected but not documented.",
                    "issue_category": None,
                    "status": "Pending",
                    "uploaded_pdf": None,
                    "drawing_file": None,
                    "drawing_number": None,
                    "equipment_type": None,
                    "financial_impact": 0.0
                })

        # ✅ Step 4: Convert manual NCRs to same format
        submitted_ncrs = [{
            "source": "Manual",
            "ncr_number": ncr.ncr_number,
            "job_number": ncr.job_number,
            "customer_name": ncr.customer_name,
            "work_order": ncr.work_order,
            "operation_number": ncr.operation_number,
            "part_name": ncr.part_name,
            "planned_hours": ncr.planned_hours,
            "actual_hours": ncr.actual_hours,
            "issue_description": ncr.issue_description,
            "issue_category": ncr.issue_category,
            "status": ncr.status,
            "uploaded_pdf": ncr.uploaded_pdf,
            "drawing_file": ncr.drawing_file,
            "drawing_number": ncr.drawing_number,
            "equipment_type": ncr.equipment_type,
            "financial_impact": ncr.financial_impact
        } for ncr in manual_ncrs]

        # ✅ Step 5: Combine and return all
        all_ncrs = auto_ncrs + submitted_ncrs
        logging.info(f"📦 Total NCRs returned: {len(all_ncrs)} (Auto: {len(auto_ncrs)}, Manual: {len(submitted_ncrs)})")

        return jsonify(all_ncrs)

    except Exception as e:
        logging.error(f"❌ Error fetching NCR data: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500



# File upload directory
NCR_FOLDER = 'static/uploads/ncr_files'
if not os.path.exists(NCR_FOLDER):
    os.makedirs(NCR_FOLDER)

@app.route('/api/ncr_report', methods=['POST'])
def submit_ncr():
    """Allow users to submit an NCR report with file uploads and mark it as Submitted."""
    try:
        data = request.form
        ncr_file = request.files.get("ncrFile")
        drawing_file = request.files.get("drawingFile")

        # ✅ File Handling
        ncr_file_path = None
        if ncr_file and ncr_file.filename:
            filename = secure_filename(ncr_file.filename)
            ncr_file_path = os.path.join(NCR_FOLDER, filename)
            ncr_file.save(ncr_file_path)

        drawing_file_path = None
        if drawing_file and drawing_file.filename:
            filename = secure_filename(drawing_file.filename)
            drawing_file_path = os.path.join(NCR_FOLDER, filename)
            drawing_file.save(drawing_file_path)

        # ✅ Create New NCR Entry
        new_ncr = NCRTracker(
            ncr_number=data.get("ncr_number"),
            job_number=data.get("job_number"),
            work_order=data.get("work_order"),
            operation_number=int(data.get("operation_number")),
            part_name=data.get("part_name"),
            customer_name=data.get("customer_name"),
            equipment_type=data.get("equipment_type"),
            planned_hours=float(data.get("planned_hours", 0)),
            actual_hours=float(data.get("actual_hours", 0)),
            issue_description=data.get("issue_description"),
            issue_category=data.get("issue_category"),
            root_cause=data.get("root_cause"),
            corrective_action=data.get("corrective_action"),
            financial_impact=float(data.get("financial_impact", 0)),
            uploaded_pdf=ncr_file_path,
            drawing_file=drawing_file_path,
            drawing_number=data.get("drawing_number"),
            status="Submitted",  # ✅ Critical fix to ensure it doesn't stay "Pending"
            created_at=datetime.utcnow()
        )

        db.session.add(new_ncr)
        db.session.commit()

        return jsonify({"message": "✅ NCR report submitted successfully"}), 201

    except Exception as e:
        app.logger.error(f"❌ Error submitting NCR: {str(e)}")
        return jsonify({"error": "Internal Server Error"}), 500









# ✅ Unified Upload Route (Handles SAPDATA, WORKLOG & PURCHASEORDERS)
@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        logging.info("📂 Starting unified file upload process...")

        if 'file' not in request.files:
            logging.error("❌ No file part in request.")
            return jsonify({"error": "No file provided"}), 400

        file = request.files['file']
        filename = secure_filename(file.filename)

        if filename == '':
            logging.error("❌ No file selected.")
            return jsonify({"error": "No file selected"}), 400

        if not filename.lower().endswith('.xlsx'):
            logging.error("❌ Invalid file format (must be .xlsx).")
            return jsonify({"error": "Invalid file format. Please upload an Excel (.xlsx) file"}), 400

        # Ensure upload directory exists
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        logging.info(f"📁 Ensured upload folder: {app.config['UPLOAD_FOLDER']}")

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # ✅ Save file in chunks
        chunk_size = 8192  # 8KB chunks
        with open(filepath, 'wb') as f:
            while True:
                chunk = file.read(chunk_size)
                if not chunk:
                    break
                f.write(chunk)
        logging.info(f"✅ File '{filename}' saved successfully.")

        # ✅ Background file processing
        def process_file():
            try:
                with app.app_context():
                    df = pd.read_excel(filepath, engine='openpyxl')

                    # ✅ File type detection
                    file_type = detect_file_type(df, filename)
                    logging.info(f"📌 File '{filename}' detected as type: {file_type}")

                    # ✅ Column Standardization
                    df = standardize_columns(df, file_type)

                    # ✅ File processing based on detected type
                    if file_type == "sapdata":
                        logging.info("📊 Beginning SAPDATA processing...")
                        process_sapdata(df)
                    elif file_type == "worklog":
                        logging.info("⏳ Beginning WORKLOG processing...")
                        process_worklog(df)
                    elif file_type == "purchaseorders":
                        logging.info("📑 Beginning PURCHASE ORDERS processing...")
                        result = process_purchase_orders(filepath)
                        if result.get('status') == 'error':
                            raise ValueError(result.get('message'))
                    else:
                        raise ValueError("❌ Unknown file type after detection.")

            except Exception as e:
                logging.error(f"❌ Background processing error for '{filename}': {str(e)}")

            finally:
                if os.path.exists(filepath):
                    os.remove(filepath)
                    logging.info(f"🗑️ Temporary file '{filename}' cleaned up after processing.")

        executor.submit(process_file)

        return jsonify({
            "status": "success",
            "message": f"File '{filename}' uploaded successfully. Processing in background."
        })

    except Exception as e:
        logging.error(f"❌ Upload error: {str(e)}")
        if os.path.exists(filepath):
            os.remove(filepath)
            logging.info(f"🗑️ File '{filename}' removed due to error.")
        return jsonify({"error": str(e)}), 500

# ✅ Enhanced File Type Detection
def detect_file_type(df, filename):
    lower_filename = filename.lower()
    if "sapdata" in lower_filename:
        return "sapdata"
    elif "worklog" in lower_filename:
        return "worklog"
    elif "purchaseorder" in lower_filename or "purchase_order" in lower_filename:
        return "purchaseorders"

    columns = set(df.columns.str.strip().str.lower())
    sapdata_cols = {'order', 'oper./act.', 'oper.workcenter', 'work', 'actual work'}
    worklog_cols = {'traficcol', 'order', 'operation', 'posting date', 'start time', 'end time'}
    purchaseorder_cols = {'req. tracking number', 'purchasing document', 'item', 'purchasing group',
                          'document date', 'vendor/supplying plant', 'short text', 'order quantity',
                          'net price', 'still to be delivered (qty)', 'still to be delivered (value)', 'material'}

    if sapdata_cols.issubset(columns):
        return "sapdata"
    elif worklog_cols.issubset(columns):
        return "worklog"
    elif purchaseorder_cols.issubset(columns):
        return "purchaseorders"
    else:
        return "unknown"



# ✅ Standardize Columns
def standardize_columns(df, file_type):
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()

    if file_type == 'sapdata':
        column_mapping = {
            'sales document': 'job_number',
            'order': 'work_order_number',
            'oper./act.': 'operation_number',
            'oper.workcenter': 'work_center',
            'description': 'part_name',
            'opr. short text': 'task_description',
            'work': 'planned_hours',
            'actual work': 'actual_hours',
            'list name': 'customer_name'
        }
    elif file_type == 'worklog':
        column_mapping = {
            "traficcol": None,
            "order": "work_order",
            "operation": "operation_number",
            "posting date": "posting_date",
            "status": None,
            "start time": "start_time",
            "end time": "end_time",
            "operation short text": "operation_description",
            "total booked hrs": "total_booked_hours",
            "confirmation text": None,
            "planned hrs": "planned_hours",
            "full name": "employee_name",
            "start date": "start_date",
            "end date": "end_date",
            "booked hrs": "booked_hours",
            "sales order": "job_number"
        }
    elif file_type == 'purchaseorders':
        column_mapping = {
            'Req. Tracking Number': 'job_number',
            'Purchasing Document': 'po_number',
            'Item': 'line_item',
            'Purchasing Group': 'buyer_initials',
            'Document Date': 'document_date',
            'Vendor/supplying plant': 'vendor',
            'Short Text': 'description',
            'Order Quantity': 'order_quantity',
            'Net price': 'net_price',
            'Still to be delivered (qty)': 'pending_quantity',
            'Still to be delivered (value)': 'pending_value',
            'Material': 'material_code',
            'Deletion Indicator': 'deletion_indicator'  # Used for filtering
        }       
    else:
        raise ValueError(f"Unknown file type: {file_type}")

    df.rename(columns={k: v for k, v in column_mapping.items() if v}, inplace=True)
    logging.debug(f"✅ Columns standardized for {file_type.upper()}: {df.columns.tolist()}")
    return df




#Start Shop Lead Purchase order Page-------------------------------

@app.route('/api/purchase', methods=['GET'])
def get_purchase_orders():
    try:
        # Fetch all purchase orders
        purchase_orders = PurchaseOrder.query.all()

        # Prepare comprehensive list of POs
        orders_data = []
        for po in purchase_orders:
            orders_data.append({
                "id": po.id,
                "po_number": po.po_number,
                "line_item": po.line_item,
                "job_number": po.job_number,
                "vendor": po.vendor,
                "buyer_initials": po.buyer_initials,
                "document_date": po.document_date.strftime('%Y-%m-%d'),
                "description": po.description,
                "order_quantity": po.order_quantity,
                "net_price": po.net_price,
                "pending_quantity": po.pending_quantity,
                "pending_value": po.pending_value,
                "material_code": po.material_code,
                "status": po.status,
                "order_date": po.order_date.strftime('%Y-%m-%d'),
                "expected_delivery": po.expected_delivery.strftime('%Y-%m-%d') if po.expected_delivery else None,
                "order_details": po.order_details
            })

        # Metrics for quick view
        open_pos = len([po for po in purchase_orders if po.status == 'Open'])
        received_today_count = len([
            po for po in purchase_orders if po.expected_delivery and po.expected_delivery.date() == datetime.utcnow().date()
            and po.status == 'Delivered'
        ])
        late_deliveries = len([
            po for po in purchase_orders if po.expected_delivery and po.expected_delivery.date() < datetime.utcnow().date()
            and po.status != 'Delivered'
        ])

        # Identify POs ready to ship (linked to job data)
        ready_to_ship_pos = []
        waiting_on_vendor_pos = []

        # Fetch all jobs to cross-reference vendor operations
        jobs = Job.query.all()

        SEND_TO_VENDOR_PATTERNS = [
            r"send to vendor", r"send to coating vendor", r"send to coating", r"send for processing",
            r"send to heat treat", r"send for machining", r"send to plating"
        ]
        RECEIVE_FROM_VENDOR_PATTERNS = [
            r"receive from vendor", r"receive material from vendor", r"receive raw material",
            r"receive from coating vendor", r"receive from plating", r"receive from heat treat"
        ]

        def matches_patterns(description, patterns):
            if not description:
                return False
            description = description.strip().lower()
            return any(re.search(pattern, description, re.IGNORECASE) for pattern in patterns)

        # Loop through jobs and operations to categorize POs
        for job in jobs:
            for wo in job.work_orders:
                sorted_operations = sorted(wo.operations, key=lambda op: op.operation_number)
                current_op = next(
                    (op for op in sorted_operations if op.status in ["Not Started", "In Progress"]), None
                )

                if current_op and current_op.task_description:
                    task_desc = current_op.task_description.lower()

                    if matches_patterns(task_desc, SEND_TO_VENDOR_PATTERNS):
                        # Find matching PO for sending
                        matching_pos = [po for po in purchase_orders if po.job_number == job.job_number and po.status == 'Open']
                        ready_to_ship_pos.extend(matching_pos)

                    elif matches_patterns(task_desc, RECEIVE_FROM_VENDOR_PATTERNS):
                        matching_pos = [po for po in purchase_orders if po.job_number == job.job_number and po.status in ['In Transit', 'Open']]
                        waiting_on_vendor_pos.extend(matching_pos)

        # Deduplicate PO lists
        ready_to_ship_pos = list({po.id: po for po in ready_to_ship_pos}.values())
        waiting_on_vendor_pos = list({po.id: po for po in waiting_on_vendor_pos}.values())

        # Calculate total_value explicitly (was included earlier)
        total_value = sum(po.net_price * po.order_quantity for po in purchase_orders)

        # Include this clearly in your metrics response
        response = {
            "purchase_orders": orders_data,
            "metrics": {
                "open_pos": open_pos,
                "received_today": received_today_count,
                "late_deliveries": late_deliveries,
                "ready_to_ship": len(ready_to_ship_pos),
                "waiting_on_vendor": len(waiting_on_vendor_pos),
                "total_value": round(total_value, 2)  # <-- explicitly add this
            },
            "vendor_statuses": {
                "ready_to_ship": [{
                    "po_number": po.po_number,
                    "job_number": po.job_number,
                    "vendor": po.vendor,
                    "material": po.material_code,
                    "quantity": po.pending_quantity,
                    "expected_delivery": po.expected_delivery.strftime('%Y-%m-%d') if po.expected_delivery else None,
                    "status": po.status
                } for po in ready_to_ship_pos],

                "waiting_on_vendor": [{
                    "po_number": po.po_number,
                    "job_number": po.job_number,
                    "vendor": po.vendor,
                    "material": po.material_code,
                    "quantity": po.pending_quantity,
                    "expected_delivery": po.expected_delivery.strftime('%Y-%m-%d') if po.expected_delivery else None,
                    "status": po.status
                } for po in waiting_on_vendor_pos]
            }
        }

        return jsonify(response)

    except Exception as e:
        logging.error(f"❌ Error fetching purchase data: {str(e)}")
        return jsonify({'error': 'Failed to fetch purchase data'}), 500


def generate_delivery_timeline(purchase_orders):
    today = datetime.utcnow().date()
    end_date = today + timedelta(days=30)
    timeline_dates = [(today + timedelta(days=i)) for i in range(31)]
    date_counter = Counter()

    for po in purchase_orders:
        if po.expected_delivery and today <= po.expected_delivery.date() <= end_date:
            date_counter[po.expected_delivery.date()] += 1

    dates = [date.strftime('%Y-%m-%d') for date in timeline_dates]
    quantities = [date_counter.get(date, 0) for date in timeline_dates]

    return {'dates': dates, 'quantities': quantities}





























#Begin Manual Upload Work History Analysis--------------------------------------------
workhistory_api = Blueprint('workhistory_api', __name__)

BURDEN_RATE = 199
@main_bp.route("/workhistory")
def workhistory_dashboard():
    return render_template("work_history.html")

#takes user to page specific to data by year
@workhistory_api.route("/workhistory/year/<int:year>")
def work_history_year(year):
    return render_template("work_history_year.html", year=year)

#This new API route works with the one above, This code is for year specific page
@workhistory_api.route("/api/workhistory/summary/year/<int:year>")
def get_yearly_summary_breakdown(year):
    from sqlalchemy import case, distinct

    BURDEN_RATE = 199

    overrun_case = case(
        (JobHistory.actual_hours > JobHistory.planned_hours,
         JobHistory.actual_hours - JobHistory.planned_hours),
        else_=0
    )
    ghost_case = case(
        (JobHistory.actual_hours == 0, JobHistory.planned_hours),
        else_=0
    )

    # 🔹 1. Summary Totals
    summary_result = db.session.query(
        func.sum(JobHistory.planned_hours),
        func.sum(JobHistory.actual_hours),
        func.sum(overrun_case),
        func.sum(JobHistory.actual_hours * BURDEN_RATE),
        func.sum(JobHistory.planned_hours * BURDEN_RATE),
        func.count(JobHistory.id),
        func.count(func.distinct(JobHistory.job_number)),
        func.count(func.distinct(JobHistory.customer_name)),
        func.sum(case((JobHistory.work_center.ilike("NCR"), JobHistory.actual_hours), else_=0)),
        func.count(func.distinct(JobHistory.part_name)),
        func.sum(ghost_case)
    ).filter(func.extract('year', JobHistory.operation_finish_date) == year).first()

    total_planned = float(summary_result[0] or 0)
    total_actual = float(summary_result[1] or 0)
    total_overrun = float(summary_result[2] or 0)
    ghost_hours = float(summary_result[10] or 0)

    opportunity_hours = total_overrun + ghost_hours
    buffer_percent = (opportunity_hours / total_planned * 100) if total_planned else 0

    summary = {
        "year": year,
        "total_planned_hours": total_planned,
        "total_actual_hours": total_actual,
        "total_overrun_hours": total_overrun,
        "ghost_hours": ghost_hours,
        "opportunity_cost_hours": opportunity_hours,
        "opportunity_cost_dollars": opportunity_hours * BURDEN_RATE,
        "recommended_buffer_percent": round(buffer_percent, 2),
        "total_actual_cost": float(summary_result[3] or 0),
        "total_planned_cost": float(summary_result[4] or 0),
        "total_operations": int(summary_result[5] or 0),
        "total_jobs": int(summary_result[6] or 0),
        "total_customers": int(summary_result[7] or 0),
        "total_ncr_hours": float(summary_result[8] or 0),
        "total_unique_parts": int(summary_result[9] or 0)
    }

    # 🔹 2. Top Overruns
    top_overruns_query = db.session.query(
        JobHistory.job_number,
        JobHistory.part_name,
        JobHistory.work_center,
        JobHistory.task_description,
        JobHistory.planned_hours,
        JobHistory.actual_hours,
        (JobHistory.actual_hours - JobHistory.planned_hours).label("overrun_hours"),
        ((JobHistory.actual_hours - JobHistory.planned_hours) * BURDEN_RATE).label("overrun_cost")
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year,
        JobHistory.actual_hours > JobHistory.planned_hours,
        ~JobHistory.task_description.ilike("%Dismantling & Inspection%")
    ).order_by(((JobHistory.actual_hours - JobHistory.planned_hours) * BURDEN_RATE).desc()).limit(10).all()

    top_overruns = [
        {
            "job_number": row.job_number,
            "part_name": row.part_name,
            "work_center": row.work_center,
            "task_description": row.task_description,
            "planned_hours": float(row.planned_hours or 0),
            "actual_hours": float(row.actual_hours or 0),
            "overrun_hours": float(row.overrun_hours or 0),
            "overrun_cost": float(row.overrun_cost or 0)
        }
        for row in top_overruns_query
    ]

    # 🔹 3. NCR Summary by Part
    ncr_summary_query = db.session.query(
        JobHistory.part_name,
        func.sum(JobHistory.actual_hours).label("total_ncr_hours"),
        func.sum(JobHistory.actual_hours * BURDEN_RATE).label("total_ncr_cost"),
        func.count(JobHistory.id).label("ncr_occurrences")
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year,
        JobHistory.work_center.ilike("NCR")
    ).group_by(JobHistory.part_name).order_by(func.sum(JobHistory.actual_hours * BURDEN_RATE).desc()).all()

    ncr_summary = [
        {
            "part_name": row.part_name,
            "total_ncr_hours": float(row.total_ncr_hours or 0),
            "total_ncr_cost": float(row.total_ncr_cost or 0),
            "ncr_occurrences": row.ncr_occurrences
        }
        for row in ncr_summary_query
    ]

    # 🔹 4. Work Center Performance
    wc_query = db.session.query(
        JobHistory.work_center,
        func.sum(JobHistory.planned_hours),
        func.sum(JobHistory.actual_hours),
        func.sum(overrun_case),
        func.sum(overrun_case * BURDEN_RATE)
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year
    ).group_by(JobHistory.work_center).order_by(func.sum(JobHistory.actual_hours).desc()).all()

    workcenter_summary = [
        {
            "work_center": row[0],
            "planned_hours": float(row[1] or 0),
            "actual_hours": float(row[2] or 0),
            "overrun_hours": float(row[3] or 0),
            "overrun_cost": float(row[4] or 0)
        }
        for row in wc_query
    ]

    # 🔹 5. Repeat NCR Failures
    repeat_ncr_raw = db.session.query(
        JobHistory.part_name,
        func.count(func.distinct(JobHistory.job_number)).label("distinct_jobs"),
        func.sum(JobHistory.actual_hours).label("repeat_ncr_hours")
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year,
        JobHistory.work_center.ilike("NCR")
    ).group_by(JobHistory.part_name).having(func.count(func.distinct(JobHistory.job_number)) > 1).all()

    repeat_ncr_failures = [
        {
            "part_name": row.part_name,
            "repeat_ncr_hours": float(row.repeat_ncr_hours or 0),
            "total_ncr_jobs": row.distinct_jobs
        }
        for row in repeat_ncr_raw
    ]

    # 🔹 6. Quarterly Summary
    quarterly_raw = db.session.query(
        JobHistory.operation_finish_date,
        func.sum(JobHistory.planned_hours),
        func.sum(JobHistory.actual_hours),
        func.sum(overrun_case),
        func.sum(overrun_case * BURDEN_RATE),
        func.count(func.distinct(JobHistory.job_number))
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year
    ).group_by(JobHistory.operation_finish_date).all()

    quarter_map = {}
    for row in quarterly_raw:
        date = row[0]
        if not date:
            continue
        q = (date.month - 1) // 3 + 1
        label = f"Q{q} {date.year}"
        if label not in quarter_map:
            quarter_map[label] = {
                "planned_hours": 0.0,
                "actual_hours": 0.0,
                "overrun_hours": 0.0,
                "overrun_cost": 0.0,
                "total_jobs": 0
            }
        quarter_map[label]["planned_hours"] += float(row[1] or 0)
        quarter_map[label]["actual_hours"] += float(row[2] or 0)
        quarter_map[label]["overrun_hours"] += float(row[3] or 0)
        quarter_map[label]["overrun_cost"] += float(row[4] or 0)
        quarter_map[label]["total_jobs"] += int(row[5] or 0)

    quarterly_summary = [
        {"quarter": label, **values}
        for label, values in sorted(quarter_map.items())
    ]

    # 🔹 7. Job Adjustments
    job_adjustments_raw = db.session.query(
        JobHistory.job_number,
        func.sum(JobHistory.planned_hours),
        func.sum(JobHistory.actual_hours),
        func.sum(overrun_case)
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year
    ).group_by(JobHistory.job_number).having(func.sum(overrun_case) > 0).all()

    job_adjustments = [
        {
            "job_number": row[0],
            "total_planned": float(row[1] or 0),
            "total_actual": float(row[2] or 0),
            "needed_increase": float(row[3] or 0)
        }
        for row in job_adjustments_raw
    ]

    # 🔹 8. Overrun Adjustment Recommendations by Part
    part_overruns_raw = db.session.query(
        JobHistory.part_name,
        func.sum(JobHistory.planned_hours).label("total_planned"),
        func.sum(JobHistory.actual_hours).label("total_actual"),
        func.sum(overrun_case).label("total_overrun")
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year,
        JobHistory.actual_hours > JobHistory.planned_hours
    ).group_by(JobHistory.part_name).having(func.sum(overrun_case) > 0).order_by(func.sum(overrun_case).desc()).limit(20).all()

    part_overruns = [
        {
            "part_name": row.part_name,
            "total_planned": float(row.total_planned or 0),
            "total_actual": float(row.total_actual or 0),
            "overrun_hours": float(row.total_overrun or 0),
            "suggested_percent_increase": round((row.total_overrun / row.total_planned) * 100, 1) if row.total_planned else 0
        }
        for row in part_overruns_raw
    ]

    # 🔹 9. Task-Level Overrun Breakdown (for parts above)
    tracked_parts = [row.part_name for row in part_overruns_raw]

    task_breakdown_raw = db.session.query(
        JobHistory.part_name,
        JobHistory.task_description,
        func.sum(JobHistory.planned_hours).label("total_planned"),
        func.sum(JobHistory.actual_hours).label("total_actual"),
        func.sum(overrun_case).label("total_overrun")
    ).filter(
        func.extract('year', JobHistory.operation_finish_date) == year,
        JobHistory.actual_hours > JobHistory.planned_hours,
        JobHistory.part_name.in_(tracked_parts)
    ).group_by(JobHistory.part_name, JobHistory.task_description).having(func.sum(overrun_case) > 0).all()

    part_task_details = [
        {
            "part_name": row.part_name,
            "task_description": row.task_description,
            "total_planned": float(row.total_planned or 0),
            "total_actual": float(row.total_actual or 0),
            "overrun_hours": float(row.total_overrun or 0),
            "suggested_percent_increase": round((row.total_overrun / row.total_planned) * 100, 1) if row.total_planned else 0
        }
        for row in task_breakdown_raw
    ]


    # 🔹 9. NCR Averages (All-Time)
    years_with_ncr = db.session.query(
        func.extract('year', JobHistory.operation_finish_date)
    ).filter(
        JobHistory.work_center.ilike("NCR")
    ).distinct().all()
    year_count = len(years_with_ncr)

    total_ncr_cost = db.session.query(
        func.sum(JobHistory.actual_hours * BURDEN_RATE)
    ).filter(
        JobHistory.work_center.ilike("NCR")
    ).scalar() or 0

    total_parts = db.session.query(
        func.count(distinct(JobHistory.part_name))
    ).filter(
        JobHistory.work_center.ilike("NCR")
    ).scalar() or 0

    ncr_averages = {
        "avg_ncr_cost_per_year": round(total_ncr_cost / year_count, 2) if year_count else 0,
        "avg_parts_with_ncr_per_year": round(total_parts / year_count, 1) if year_count else 0
    }

    # ✅ Final JSON Response
    return jsonify({
        "summary": summary,
        "top_overruns": top_overruns,
        "ncr_summary": ncr_summary,
        "workcenter_summary": workcenter_summary,
        "repeat_ncr_failures": repeat_ncr_failures,
        "quarterly_summary": quarterly_summary,
        "job_adjustments": job_adjustments,
        "part_overruns": part_overruns,
        "part_task_details": part_task_details,
        "ncr_averages": ncr_averages
    })



#API for fetching job specific information when clicking a part in NCR section
@workhistory_api.route("/api/workhistory/ncr/details")
def get_ncr_part_details():
    from sqlalchemy import distinct

    year = request.args.get("year", type=int)
    part = request.args.get("part", type=str)

    # 🔹 1. Specific Part Breakdown (year + part filter)
    part_results = db.session.query(
        JobHistory.job_number,
        JobHistory.work_order_number,
        func.sum(JobHistory.actual_hours).label("ncr_hours")
    ).filter(
        JobHistory.part_name == part,
        JobHistory.work_center.ilike("NCR"),
        func.extract('year', JobHistory.operation_finish_date) == year
    ).group_by(JobHistory.job_number, JobHistory.work_order_number).all()

    # 🔹 2. All-Time Summary Stats
    # Get distinct years with NCR activity
    years_with_ncr = db.session.query(
        func.extract('year', JobHistory.operation_finish_date).label("yr")
    ).filter(
        JobHistory.work_center.ilike("NCR")
    ).distinct().all()
    year_count = len(years_with_ncr)

    # Total NCR cost and part count across all years
    total_ncr_cost = db.session.query(
        func.sum(JobHistory.actual_hours * 199)
    ).filter(
        JobHistory.work_center.ilike("NCR")
    ).scalar() or 0

    total_parts = db.session.query(
        func.count(distinct(JobHistory.part_name))
    ).filter(
        JobHistory.work_center.ilike("NCR")
    ).scalar() or 0

    # 🔹 3. Safe averages
    avg_cost_per_year = total_ncr_cost / year_count if year_count else 0
    avg_parts_per_year = total_parts / year_count if year_count else 0

    return jsonify({
        "job_data": [
            {
                "job_number": row.job_number,
                "work_order_number": row.work_order_number,
                "ncr_hours": float(row.ncr_hours or 0)
            } for row in part_results
        ],
        "all_time_averages": {
            "avg_ncr_cost_per_year": round(avg_cost_per_year, 2),
            "avg_parts_with_ncr_per_year": round(avg_parts_per_year, 1)
        }
    })


# 1. Yearly Summary - Main page
@workhistory_api.route("/api/workhistory/summary/yearly")
def get_yearly_summary():
    results = db.session.query(
        func.extract('year', JobHistory.operation_finish_date).label("year"),
        func.count(func.distinct(JobHistory.work_order_number)).label("work_orders"),
        func.count(func.distinct(JobHistory.part_name)).label("unique_parts"),
        func.sum(JobHistory.planned_hours).label("planned_hours"),
        func.sum(JobHistory.actual_hours).label("actual_hours"),
        (func.sum(JobHistory.actual_hours) * BURDEN_RATE).label("actual_cost")
    ).group_by("year").order_by("year").all()

    return jsonify([dict(row._asdict()) for row in results])

@workhistory_api.route("/api/workhistory/summary/full")
def get_full_summary():
    from sqlalchemy import case
    import logging

    logger = logging.getLogger(__name__)
    BURDEN_RATE = 199

    try:
        # 🔍 Print sample records
        sample_rows = db.session.query(
            JobHistory.job_number,
            JobHistory.task_description,
            JobHistory.planned_hours,
            JobHistory.actual_hours,
            JobHistory.operation_finish_date,
            JobHistory.recorded_date
        ).limit(5).all()

        for row in sample_rows:
            logger.info(f"🔍 Sample Job: {row}")

        # ✅ Safe CASE syntax for SQLAlchemy 2.x
        overrun_case = case(
            (JobHistory.actual_hours > JobHistory.planned_hours,
             JobHistory.actual_hours - JobHistory.planned_hours),
            else_=0
        )

        ncr_case = case(
            (JobHistory.work_center == "NCR", JobHistory.actual_hours),
            else_=0
        )


        # --- Summary Metrics ---
        summary_result = db.session.query(
            func.sum(JobHistory.planned_hours),
            func.sum(JobHistory.actual_hours),
            func.sum(overrun_case),
            func.sum(JobHistory.actual_hours * BURDEN_RATE),
            func.sum(JobHistory.planned_hours * BURDEN_RATE),
            func.count(JobHistory.id),
            func.count(func.distinct(JobHistory.job_number)),
            func.count(func.distinct(JobHistory.customer_name)),
            func.sum(ncr_case),
            func.count(func.distinct(JobHistory.part_name))
        ).first()

        if not summary_result:
            logger.warning("No data returned in summary query.")
            return jsonify({"error": "No data found"}), 404

        summary = {
            "total_planned_hours": float(summary_result[0] or 0),
            "total_actual_hours": float(summary_result[1] or 0),
            "total_overrun_hours": float(summary_result[2] or 0),
            "total_actual_cost": float(summary_result[3] or 0),
            "total_planned_cost": float(summary_result[4] or 0),
            "total_operations": int(summary_result[5] or 0),
            "total_jobs": int(summary_result[6] or 0),
            "total_customers": int(summary_result[7] or 0),
            "total_ncr_hours": float(summary_result[8] or 0),
            "total_unique_parts": int(summary_result[9] or 0)
        }

        # --- Yearly Breakdown ---
        yearly_query = db.session.query(
            func.extract('year', JobHistory.operation_finish_date).label("year"),
            func.sum(JobHistory.planned_hours),
            func.sum(JobHistory.actual_hours),
            func.sum(overrun_case),
            func.sum(ncr_case),
            func.count(func.distinct(JobHistory.job_number)),
            func.count(JobHistory.id),
            func.count(func.distinct(JobHistory.customer_name))
        ).group_by("year").order_by("year").all()

        yearly_breakdown = []
        for row in yearly_query:
            try:
                yearly_breakdown.append({
                    "year": int(row[0]) if row[0] else None,
                    "planned_hours": float(row[1] or 0),
                    "actual_hours": float(row[2] or 0),
                    "overrun_hours": float(row[3] or 0),
                    "ncr_hours": float(row[4] or 0),
                    "job_count": int(row[5] or 0),
                    "operation_count": int(row[6] or 0),
                    "customer_count": int(row[7] or 0)
                })
            except Exception as e:
                logger.warning(f"⚠️ Failed to parse yearly row: {row} → {e}")

        # --- Work Center Breakdown ---
        wc_query = db.session.query(
            JobHistory.work_center,
            func.sum(JobHistory.planned_hours),
            func.sum(JobHistory.actual_hours),
            func.sum(overrun_case)
        ).group_by(JobHistory.work_center).order_by(func.sum(JobHistory.actual_hours).desc()).all()

        workcenter_breakdown = []
        for row in wc_query:
            try:
                workcenter_breakdown.append({
                    "work_center": row[0] or "UNKNOWN",
                    "total_planned_hours": float(row[1] or 0),
                    "total_actual_hours": float(row[2] or 0),
                    "overrun_hours": float(row[3] or 0)
                })
            except Exception as e:
                logger.warning(f"⚠️ Failed to parse work center row: {row} → {e}")

        logger.info("✅ Full summary API returned successfully.")

        return jsonify({
            "summary": summary,
            "yearly_breakdown": yearly_breakdown,
            "workcenter_breakdown": workcenter_breakdown
        })

    except Exception as e:
        logger.exception("❌ Error in /summary/full API")
        return jsonify({"error": "Internal server error", "details": str(e)}), 500




# 2. Customer Profitability Summary
@workhistory_api.route("/api/workhistory/summary/customers")
def get_customer_summary():
    results = db.session.query(
        JobHistory.customer_name,
        func.sum(JobHistory.planned_hours).label("planned_hours"),
        func.sum(JobHistory.actual_hours).label("actual_hours"),
        ((func.sum(JobHistory.planned_hours) - func.sum(JobHistory.actual_hours)) * BURDEN_RATE).label("profit_loss")
    ).group_by(JobHistory.customer_name).order_by(func.sum(JobHistory.actual_hours).desc()).all()

    return jsonify([dict(row._asdict()) for row in results])


# 3. Part Performance Summary
@workhistory_api.route("/api/workhistory/summary/parts")
def get_part_summary():
    results = db.session.query(
        JobHistory.part_name,
        func.count(func.distinct(JobHistory.work_order_number)).label("job_count"),
        func.sum(JobHistory.planned_hours).label("planned_hours"),
        func.sum(JobHistory.actual_hours).label("actual_hours"),
        ((func.sum(JobHistory.planned_hours) - func.sum(JobHistory.actual_hours)) * BURDEN_RATE).label("roi")
    ).group_by(JobHistory.part_name).order_by(func.sum(JobHistory.actual_hours).desc()).limit(100).all()

    return jsonify([dict(row._asdict()) for row in results])


# 4. Work Center Trends
@workhistory_api.route("/api/workhistory/summary/workcenters")
def get_workcenter_summary():
    results = db.session.query(
        JobHistory.work_center,
        func.count(JobHistory.operation_number).label("operations"),
        func.sum(JobHistory.planned_hours).label("planned_hours"),
        func.sum(JobHistory.actual_hours).label("actual_hours"),
        ((func.sum(JobHistory.actual_hours) - func.sum(JobHistory.planned_hours)) * BURDEN_RATE).label("overrun_cost")
    ).group_by(JobHistory.work_center).order_by(func.sum(JobHistory.actual_hours).desc()).all()

    return jsonify([dict(row._asdict()) for row in results])


# 5. Deep Dive Filtering
@workhistory_api.route("/api/workhistory/filter")
def filter_work_history():
    year = request.args.get("year")
    customer = request.args.get("customer")
    part = request.args.get("part")
    work_center = request.args.get("work_center")

    query = db.session.query(JobHistory)

    if year:
        query = query.filter(func.extract('year', JobHistory.operation_finish_date) == int(year))
    if customer:
        query = query.filter(JobHistory.customer_name.ilike(f"%{customer}%"))
    if part:
        query = query.filter(JobHistory.part_name.ilike(f"%{part}%"))
    if work_center:
        query = query.filter(JobHistory.work_center.ilike(f"%{work_center}%"))

    records = query.limit(1000).all()

    return jsonify([
        {
            "job_number": r.job_number,
            "customer_name": r.customer_name,
            "part_name": r.part_name,
            "work_center": r.work_center,
            "planned_hours": r.planned_hours,
            "actual_hours": r.actual_hours,
            "operation_finish_date": r.operation_finish_date,
        }
        for r in records
    ])


# 6. Trend Analysis (rolling yearly cost)
@workhistory_api.route("/api/workhistory/trends")
def get_trends():
    results = db.session.query(
        func.extract('year', JobHistory.operation_finish_date).label("year"),
        func.sum(JobHistory.actual_hours * BURDEN_RATE).label("total_cost")
    ).group_by("year").order_by("year").all()

    return jsonify([dict(row._asdict()) for row in results])














#Work history clickable cards to take user to metrics page
@workhistory_api.route("/workhistory/metric/<metric>")
def view_metric_detail_page(metric):
    return render_template("metric_detail.html", metric=metric)

@workhistory_api.route("/api/workhistory/metric/<metric>")
def get_metric_detail(metric):
    from sqlalchemy import or_
    import logging

    logger = logging.getLogger(__name__)
    BURDEN_RATE = 199

    valid_metrics = {
        "ncr_hours", "planned_hours", "actual_hours", "overrun_hours",
        "planned_cost", "actual_cost", "overrun_cost", "overrun_percent",
        "total_operations", "total_jobs", "total_customers",
        "avg_cost_per_hour", "avg_job_size"
    }

    if metric not in valid_metrics:
        return jsonify({"error": f"Unsupported metric: {metric}"}), 400

    try:
        filters = {
            "ncr_hours": JobHistory.work_center == "NCR",
            "planned_hours": JobHistory.planned_hours > 0,
            "actual_hours": JobHistory.actual_hours > 0,
            "planned_cost": None,
            "actual_cost": None,
            "overrun_hours": JobHistory.actual_hours > JobHistory.planned_hours,
            "overrun_cost": JobHistory.actual_hours > JobHistory.planned_hours,
            "overrun_percent": JobHistory.actual_hours > JobHistory.planned_hours,
            "avg_cost_per_hour": None,
            "avg_job_size": None,
            "total_operations": None,
            "total_jobs": None,
            "total_customers": None
        }

        filter_condition = filters[metric]

        query = db.session.query(
            JobHistory.job_number,
            JobHistory.customer_name,
            JobHistory.part_name,
            JobHistory.work_order_number,
            JobHistory.work_center,
            JobHistory.task_description,
            JobHistory.planned_hours,
            JobHistory.actual_hours,
            JobHistory.operation_finish_date
        )

        if filter_condition is not None:
            query = query.filter(filter_condition)

        results = query.order_by(JobHistory.operation_finish_date.desc()).all()

        rows = []
        for row in results:
            overrun = (row.actual_hours or 0) - (row.planned_hours or 0)
            rows.append({
                "job_number": row.job_number,
                "customer_name": row.customer_name,
                "part_name": row.part_name,
                "work_order_number": row.work_order_number,
                "work_center": row.work_center,
                "task_description": row.task_description,
                "planned_hours": float(row.planned_hours or 0),
                "actual_hours": float(row.actual_hours or 0),
                "overrun_hours": float(overrun if overrun > 0 else 0),
                "finish_date": row.operation_finish_date.strftime("%Y-%m-%d") if row.operation_finish_date else ""
            })

        logger.info(f"🔍 Metric '{metric}' returned {len(rows)} rows")
        return jsonify({
            "metric": metric,
            "count": len(rows),
            "rows": rows
        })

    except Exception as e:
        logger.exception(f"❌ Error in metric detail API for '{metric}'")
        return jsonify({"error": "Internal server error", "details": str(e)}), 500







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