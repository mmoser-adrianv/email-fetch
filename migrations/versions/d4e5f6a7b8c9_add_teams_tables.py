"""add_teams_tables

Revision ID: d4e5f6a7b8c9
Revises: c3a6b2a63242
Create Date: 2026-03-09 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'd4e5f6a7b8c9'
down_revision = 'c3a6b2a63242'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'teams_chats',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('chat_id', sa.String(length=512), nullable=False),
        sa.Column('display_name', sa.String(length=512), nullable=True),
        sa.Column('chat_type', sa.String(length=64), nullable=True),
        sa.Column('last_updated_date_time', sa.DateTime(timezone=True), nullable=True),
        sa.Column('member_count', sa.Integer(), nullable=True),
        sa.Column('scraped_at', sa.DateTime(), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('chat_id'),
    )
    op.create_index('ix_teams_chats_chat_id', 'teams_chats', ['chat_id'], unique=True)

    op.create_table(
        'teams_messages',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('teams_chat_id', sa.Integer(), nullable=False),
        sa.Column('message_id', sa.String(length=512), nullable=False),
        sa.Column('sender_name', sa.String(length=255), nullable=True),
        sa.Column('sender_email', sa.String(length=255), nullable=True),
        sa.Column('content_html', sa.Text(), nullable=True),
        sa.Column('content_text', sa.Text(), nullable=True),
        sa.Column('created_date_time', sa.DateTime(timezone=True), nullable=True),
        sa.Column('message_type', sa.String(length=64), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.ForeignKeyConstraint(['teams_chat_id'], ['teams_chats.id'], ),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('message_id'),
    )
    op.create_index('ix_teams_messages_message_id', 'teams_messages', ['message_id'], unique=True)


def downgrade():
    op.drop_index('ix_teams_messages_message_id', table_name='teams_messages')
    op.drop_table('teams_messages')
    op.drop_index('ix_teams_chats_chat_id', table_name='teams_chats')
    op.drop_table('teams_chats')
