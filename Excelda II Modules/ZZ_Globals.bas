Option Explicit
'Centralized globals for the project.
'Moved here from multiple modules to improve maintainability.
'
'Any of these can be refined to explicit types later. For now use Variant to preserve behavior.

'Link / Game state
Public currentScreen As Variant
Public screenSetUpTimer As Variant
Public linkCellAddress As Variant
Public CodeCell As Variant
Public moveDir As Variant
Public lastDir As Variant
Public LinkSprite As Variant
Public gameSpeed As Variant
Public LinkSpriteTop As Variant
Public LinkSpriteLeft As Variant
Public LinkMove As Variant
Public LinkSpriteFrame As Variant

'Input / actions
Public CItem As Variant
Public DItem As Variant
Public CPress As Variant
Public DPress As Variant

'Sprites
Public SwordFrame1 As Variant
Public SwordFrame2 As Variant
Public SwordFrame3 As Variant
Public shieldSprite As Variant

'Enemy globals (four slots)
Public RNDenemyName1 As Variant, RNDenemyFrame1_1 As Variant, RNDenemyFrame1_2 As Variant
Public RNDenemyInitialCount1 As Variant, RNDenemyCount1 As Variant
Public RNDenemyDir1 As Variant, RNDenemySpeed1 As Variant, RNDenemyBehaviour1 As Variant, RNDenemyChangeRotation1 As Variant
Public RNDenemyCanShoot1 As Variant, RNDenemyChargeSpeed1 As Variant
Public RNDenemyCanCollide1 As Variant, RNDenemyCollisionDamage1 As Variant, RNDenemyShootDamage1 As Variant, RNDenemyChargeDamage1 As Variant
Public RNDenemyHit1 As Variant, RNDenemyLife1 As Variant

Public RNDenemyName2 As Variant, RNDenemyFrame2_1 As Variant, RNDenemyFrame2_2 As Variant
Public RNDenemyInitialCount2 As Variant, RNDenemyCount2 As Variant
Public RNDenemyDir2 As Variant, RNDenemySpeed2 As Variant, RNDenemyBehaviour2 As Variant, RNDenemyChangeRotation2 As Variant
Public RNDenemyCanShoot2 As Variant, RNDenemyChargeSpeed2 As Variant
Public RNDenemyCanCollide2 As Variant, RNDenemyCollisionDamage2 As Variant, RNDenemyShootDamage2 As Variant, RNDenemyChargeDamage2 As Variant
Public RNDenemyHit2 As Variant, RNDenemyLife2 As Variant

Public RNDenemyName3 As Variant, RNDenemyFrame3_1 As Variant, RNDenemyFrame3_2 As Variant
Public RNDenemyInitialCount3 As Variant, RNDenemyCount3 As Variant
Public RNDenemyDir3 As Variant, RNDenemySpeed3 As Variant, RNDenemyBehaviour3 As Variant, RNDenemyChangeRotation3 As Variant
Public RNDenemyCanShoot3 As Variant, RNDenemyChargeSpeed3 As Variant
Public RNDenemyCanCollide3 As Variant, RNDenemyCollisionDamage3 As Variant, RNDenemyShootDamage3 As Variant, RNDenemyChargeDamage3 As Variant
Public RNDenemyHit3 As Variant, RNDenemyLife3 As Variant

Public RNDenemyName4 As Variant, RNDenemyFrame4_1 As Variant, RNDenemyFrame4_2 As Variant
Public RNDenemyInitialCount4 As Variant, RNDenemyCount4 As Variant
Public RNDenemyDir4 As Variant, RNDenemySpeed4 As Variant, RNDenemyBehaviour4 As Variant, RNDenemyChangeRotation4 As Variant
Public RNDenemyCanShoot4 As Variant, RNDenemyChargeSpeed4 As Variant
Public RNDenemyCanCollide4 As Variant, RNDenemyCollisionDamage4 As Variant, RNDenemyShootDamage4 As Variant, RNDenemyChargeDamage4 As Variant
Public RNDenemyHit4 As Variant, RNDenemyLife4 As Variant

'Projectiles
Public projectileName1 As Variant, projectileSpeed1 As Variant, projectileBehaviour1 As Variant, projectileDir1 As Variant

'Triggers / collision state
Public TriggerCel As Variant
Public CollidedWith As Variant
Public RNDBounceback As Variant
Public SwordHit As Variant
Public RNDEnemyBounceback1 As Variant, RNDEnemyBounceback2 As Variant, RNDEnemyBounceback3 As Variant, RNDEnemyBounceback4 As Variant
Public UseLegacyTick As Boolean
