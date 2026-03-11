"""add_calendar_tables

Revision ID: a1b2c3d4e5f7
Revises: f6a7b8c9d0e1
Create Date: 2026-03-10 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'a1b2c3d4e5f7'
down_revision = 'f6a7b8c9d0e1'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'calendar_events',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('event_id', sa.String(length=512), nullable=False),
        sa.Column('subject', sa.Text(), nullable=True),
        sa.Column('organizer_email', sa.String(length=255), nullable=True),
        sa.Column('organizer_name', sa.String(length=255), nullable=True),
        sa.Column('start_datetime', sa.DateTime(timezone=True), nullable=True),
        sa.Column('end_datetime', sa.DateTime(timezone=True), nullable=True),
        sa.Column('timezone', sa.String(length=64), nullable=True),
        sa.Column('location', sa.Text(), nullable=True),
        sa.Column('body_html', sa.Text(), nullable=True),
        sa.Column('body_text', sa.Text(), nullable=True),
        sa.Column('is_online_meeting', sa.Boolean(), nullable=True),
        sa.Column('online_meeting_url', sa.Text(), nullable=True),
        sa.Column('join_url', sa.Text(), nullable=True),
        sa.Column('web_link', sa.Text(), nullable=True),
        sa.Column('searched_email', sa.String(length=255), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('event_id'),
    )
    op.create_index('ix_calendar_events_event_id', 'calendar_events', ['event_id'])
    op.create_index('ix_calendar_events_searched_email', 'calendar_events', ['searched_email'])

    op.create_table(
        'calendar_event_attendees',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('calendar_event_id', sa.Integer(), nullable=False),
        sa.Column('email', sa.String(length=255), nullable=True),
        sa.Column('name', sa.String(length=255), nullable=True),
        sa.Column('attendee_type', sa.String(length=32), nullable=True),
        sa.Column('response_status', sa.String(length=32), nullable=True),
        sa.ForeignKeyConstraint(['calendar_event_id'], ['calendar_events.id']),
        sa.PrimaryKeyConstraint('id'),
    )


def downgrade():
    op.drop_table('calendar_event_attendees')
    op.drop_index('ix_calendar_events_searched_email', table_name='calendar_events')
    op.drop_index('ix_calendar_events_event_id', table_name='calendar_events')
    op.drop_table('calendar_events')
