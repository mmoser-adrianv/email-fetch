"""add chat_sessions and chat_messages tables

Revision ID: b2c3d4e5f6a7
Revises: a1b2c3d4e5f6
Create Date: 2026-03-06 13:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'b2c3d4e5f6a7'
down_revision = 'a1b2c3d4e5f6'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table('chat_sessions',
    sa.Column('id', sa.Integer(), nullable=False),
    sa.Column('user_oid', sa.String(length=64), nullable=False),
    sa.Column('title', sa.String(length=255), nullable=True),
    sa.Column('openai_response_id', sa.String(length=255), nullable=True),
    sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
    sa.Column('updated_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
    sa.PrimaryKeyConstraint('id')
    )
    with op.batch_alter_table('chat_sessions', schema=None) as batch_op:
        batch_op.create_index(batch_op.f('ix_chat_sessions_user_oid'), ['user_oid'], unique=False)

    op.create_table('chat_messages',
    sa.Column('id', sa.Integer(), nullable=False),
    sa.Column('session_id', sa.Integer(), nullable=False),
    sa.Column('role', sa.String(length=16), nullable=True),
    sa.Column('content', sa.Text(), nullable=True),
    sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
    sa.ForeignKeyConstraint(['session_id'], ['chat_sessions.id'], ),
    sa.PrimaryKeyConstraint('id')
    )


def downgrade():
    op.drop_table('chat_messages')
    with op.batch_alter_table('chat_sessions', schema=None) as batch_op:
        batch_op.drop_index(batch_op.f('ix_chat_sessions_user_oid'))
    op.drop_table('chat_sessions')
