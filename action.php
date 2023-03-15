<?php

if(!defined('DOKU_INC')) die();

if(!defined('DOKU_PLUGIN')) define('DOKU_PLUGIN',DOKU_INC.'lib/plugins/');
require_once(DOKU_PLUGIN.'action.php');

/**
 * DokuWiki Plugin xlsx2dw (Action Component)
 *
 * @license GPL 2 http://www.gnu.org/licenses/gpl-2.0.html
 * @author  moevm <ydginster@gmail.com>
 */
class action_plugin_xlsx2dw extends DokuWiki_Action_Plugin 
{

    /** @inheritDoc */
    public function register(Doku_Event_Handler $controller)
    {
        $controller->register_hook('TOOLBAR_DEFINE', 'AFTER', $this, 'insert_button', array());
    }

    /**
     * @param Doku_Event $event  event object by reference
     * @param mixed      $param  optional parameter passed when event was registered
     * @return void
     */
    public function insert_button(Doku_Event $event, $param) {
        $config = require(DOKU_PLUGIN . 'xlsx2dw/conf/config.php');

        $event->data[] = [
            'type'   => $config['buttonType'],
            'title'  => $config['buttonTitle'],
            'icon'   => $config['buttonIcon'],
            'id'     => $config['buttonID'],
            'block'  => $config['buttonBlock'],
            'open'   => $config['buttonOpen'],
            'close'  => $config['buttonClose']
        ];
    }
}

